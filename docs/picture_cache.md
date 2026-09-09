# `UserPicture2`: one call, and the caches under it

`Fill.UserPicture` makes the caller responsible for everything: reading the file,
knowing whether the Shape can take a picture fill, and re-doing all of it on
every call. `BB_ApplyPicture` - `UserPicture2` in VBA - takes a Shape and a path
and decides the rest.

```vb
BlipBridge.UserPicture2 shp, "C:\textures\brick.png"
```

## The dispatch decision

```text
Shape class has a validated native path?  -> accelerated apply, cached texture
otherwise, Fill.UserPicture accepts it?   -> Office's own path
neither                                   -> BB_E_UNSUPPORTED_SHAPE, by name
```

The verdict comes from `ClassifyShapeForNativePictureFill`, the one semantic
authority, which the raw texture API asks too - so the two surfaces cannot
disagree about a Shape. Which classes are in it, and what each
one had to survive to get there, is in
[shape_compatibility.md](shape_compatibility.md).

### Falling back is about the Shape class, never about failure

This is the distinction the design turns on. The fallback exists because some
Shape classes have no native path - a Table has no receiver chain at all. It does
**not** exist to catch native failures.

So if the native path fails because the Office build is not validated, because
the Shape was deleted, or because the image will not decode, that is returned as
itself:

| Condition | Result |
|---|---|
| Office build not validated | `BB_E_UNSUPPORTED_BUILD` |
| not running in PowerPoint | `BB_E_UNSUPPORTED_HOST` |
| Shape class has no picture-fill path | `BB_E_UNSUPPORTED_SHAPE` |
| pointer is not a usable Shape | `BB_E_INVALID_SHAPE` |
| path missing, unreadable, or a directory | `BB_E_FILE_NOT_FOUND` |
| file present but not a decodable image | `BB_E_DECODE_FAILED` |
| fell back, and `Fill.UserPicture` refused too | `BB_E_FALLBACK_FAILED` |
| the fill itself failed | `BB_E_APPLY_FAILED` |

Retrying a native failure on the slower path would turn "your Office build is
unsupported" into "everything is mysteriously slow", and would hide precisely the
problems worth knowing about. `BB_GetLastError` carries the sentence-long reason
in every case.

## Two caches

### Path to texture

The first time a path is seen, the file is read and decoded into a texture; after
that the texture is reused, however many Shapes it is applied to.

The key is **path plus size plus last-write time**, not the path alone. Editing
the file on disk therefore produces the new image rather than a stale one, and
the superseded texture is released instead of accumulating. The check is one
`GetFileAttributesExW`, about two microseconds against roughly 190 for an apply.

#### Why not hash the contents

A content hash would be strictly stronger: it would catch an edit that preserved
both size and timestamp, which the current key cannot. It is deliberately not
used, and the tradeoff is worth stating rather than leaving implied.

Hashing means reading the whole file on **every** call, not just the first. That
is the one thing this cache exists to avoid: the cache turns a repeated apply
into a dictionary lookup, and re-reading a megabyte to validate the lookup would
cost far more than the 190 microseconds it saves. The metadata check reads no
file content at all.

What the current key misses is narrow. Every ordinary way of changing a file -
an editor saving, a build step writing, a copy, a download - changes the size, the
last-write time, or both. Defeating it takes a deliberate same-size write with a
restored timestamp.

So the rule is: **keep the metadata key unless a stale-cache collision is
actually observed.** If one ever is, the fix is not to hash on the hot path but to
hash at *preload* time - when a texture is first created - and compare hashes only
when the metadata says the file changed. That keeps the repeated-apply path free.
Until then, adding a hash would be paying a certain cost for a hypothetical bug.

A caller who knows a file changed underneath it in that pathological way has
`BB_ClearPictureCache`.

### Shape to last texture

Applying the same image to the same Shape twice in a row does no Office work.

The key is a composite read out of the object - the presentation's identity, the
slide's `SlideID`, and the Shape's `Id` - because neither a pointer nor an `Id`
alone would do. A pointer can be freed and its address reused; `Shape.Id` is
unique within a slide, not across a presentation and certainly not across open
documents. Reading the composite costs three Automation property fetches at about
a microsecond each.

A Shape that cannot produce a complete key is simply **not cached**. It still
gets a correct apply, it just pays for it. Inventing a key would risk showing the
wrong image, which is worse than paying for one transaction.

### Groups, and why they are never cached

A group child *can* be keyed. `tools/probe_shape_identity.ps1` measured its
`Parent` to be the **Slide**, not the group, so it reports a `SlideID`, and its
`Id` collides with nothing on the slide:

```text
child 1 : Id=5 Parent=Slide1 SlideID=256 ParentGroup=7
all ids: top:2 top:3 top:4 top:7 child:5 child:6
ids total 6, distinct 6      COLLISION: False
```

It is still not cached, for a better reason. **Filling a group changes what its
children render.** `tools/probe_group_fill_propagation.ps1` renders a child to
PNG before and after the group is filled and compares the bytes:

```text
child render before group fill: 33515 bytes
child render after  group fill: 35725 bytes
CHILD RENDER UNCHANGED BY GROUP FILL: False
```

`Fill.Type` reads 6 throughout, so no cheap property reveals the change. That
makes a child's remembered texture stale the moment its group is filled - a later
request to restore the child's own image would be skipped, and the child would
keep showing the group's picture. Silently. The reverse holds too: filling a
child changes what the group displays.

Tracking that correctly would mean invalidating a subtree in both directions on
every group and child apply, nested groups included. The cheap and obviously
correct rule is to cache neither: **a group, and anything inside one, always does
real work.** One `ParentGroup` read per apply is the whole cost of being sure,
and a Shape becomes cacheable again as soon as it is ungrouped.

This was a real defect in the first version of this cache, caught by
`tools/test_shape_lifecycle_cache.ps1`.

Nothing in either cache holds a Shape, a receiver, or any other document object.
The Shape cache stores identity *values*, so it never needs to be told when a
Shape dies.

## What the skip is worth

Measured through the real ABI, 200 iterations, same image and same Shape:

| | per call |
|---|---:|
| skipped (cache hit) | 0.49 ms |
| real apply | 6.29 ms |

Read that with its caveat. Both legs are driven across a COM boundary, and the
real-apply figure is much larger than the 0.19 ms an in-process apply costs -
because changing a fill also makes PowerPoint **repaint**, and the host gets to
pump messages between cross-process calls. The in-process profile in
[cost_profile.md](cost_profile.md) deliberately never pumps messages and so never
shows that cost.

The honest summary is that the skip avoids the apply *and* the repaint it would
trigger, which is why it is worth more in a live document than the raw apply
number suggests.

## When the cache must be invalidated

The skip verifies that `Fill.Type` is still a picture fill before trusting
itself, so these are handled automatically:

* the fill was cleared, or replaced with a solid colour or a gradient;
* the Shape was deleted - its key stops matching anything live;
* `BB_Shutdown`, `BB_ClearTextures` or `BB_ClearPictureCache` ran;
* the source file changed on disk.

These are **not** detected, because detecting them would cost more than the apply
they save:

* the fill was replaced with a *different picture* by something else, including a
  direct `Fill.UserPicture` call outside this API - it still reads as a picture
  fill;
* a Shape was copied or duplicated, and the copy inherited the fill;
* a presentation was closed and reopened, so old keys may match new Shapes that
  happen to share a `SlideID` and `Id`.

For the first, call `BB_InvalidateShape` on that Shape. For the others, call
`BB_ClearPictureCache` - after closing or reloading a document, or after
rewriting a batch of source files. It releases everything and costs nothing to
rebuild.

## Verifying

`tools/test_picture_cache.ps1` asserts all of the above in a live PowerPoint,
through the real exported functions: dispatch to each of the three routes, the
path cache, the skip, the freshness check, the solid-colour re-apply, explicit
invalidation, and one error of each kind. It fails the script on a wrong answer
rather than printing something for a human to read past.
