# Two kinds of validation, and why both are required

BlipBridge calls undocumented Office internals. Everything it will not do is as
important as everything it does, and the guards fall into two layers that answer
genuinely different questions.

```text
Structural validation
    proves that an internal object is the Office object we think it is.

Semantic validation
    proves that performing a picture fill on that Shape class is
    meaningful and safe.
```

Neither implies the other. That is not a theoretical caution - it is the lesson
of a bug that terminated PowerPoint.

## The Connector

A Connector, on the validated build:

* reports `Shape.Type = 1` - **msoAutoShape**, the same as a rectangle;
* exposes a PPCORE delegating wrapper at the same vtable, with the same inner
  offset and the same thunk shape;
* whose inner object presents the **same** OART `FillFormat` vtable,
  `oart.dll+0xAF60B8`;
* reachable through a valid control block;
* to a receiver with the **same** receiver vtable.

Every structural check passes. The read-only classifier reports the full chain.
Two Shapes, one fillable and one not, are indistinguishable by layout.

And yet:

| | result |
|---|---|
| AutoShape, native apply | survives, `Fill.Type = 6` |
| Connector, `Fill.UserPicture` | **Office itself refuses** - "value out of range" |
| Connector, native apply | **RPC 0x800706BE - the host is gone** |

The middle row explains the third. Office validates in PPCORE *before* the fill
handler is ever reached, and a connector fails that check: a connector is a line,
and a line has no interior to fill. Our native apply faithfully reproduces the
handler - and does not reproduce the pre-check - so it drives the handler into a
state Office never permits.

Reproduce it with `tools/test_connector_isolation.ps1`.

## What follows

**`Shape.Type` is descriptive metadata, not proof.** It says what PowerPoint
calls the object. It does not say what operations are valid on it. Two examples
from the same build, both reporting msoAutoShape:

| Class | reports | behaves |
|---|---|---|
| Connector | msoAutoShape | fatal under a native apply |
| WordArt | msoAutoShape | fills correctly, keeps its text and geometry |

**Structural validation stays**, and is still necessary. It is what stops the
code reading a field out of an object that is not the object it expects, and it
is what makes an unvalidated Office build fail closed. It simply cannot answer
the semantic question, because at the layout level there is nothing to see.

## Where the two layers live

| Layer | Home | Refuses |
|---|---|---|
| Build and module identity | `oart_layout.cpp` | an unvalidated Office build, a moved vtable, a patched entry point |
| Structural chain | `ResolveFillTarget` | an object that is not a PPCORE fill wrapper, or whose inner object is not the OART `FillFormat` |
| **Semantic eligibility** | `shape_policy.cpp` | a Shape class with no meaningful or safe picture fill |
| Handle validity | `native_texture.cpp` | an unknown or released texture |

`ClassifyShapeForNativePictureFill` is the **single** authority on the semantic
layer. The C ABI, the COM compatibility surface and the `UserPicture2` dispatcher
all ask it; none of them re-derives the answer. Scattering Shape-type checks was
how the two surfaces were able to disagree in the first place.

## The order the native path requires

No private Office function is called until every one of these holds:

```text
1. a valid public Shape object          (pointer, committed memory, IDispatch)
2. a supported Office build             (oart, ppcore, gfx all validated)
3. semantic eligibility                 (NativeSupported for this class)
4. a valid structural chain             (wrapper -> FillFormat -> block -> receiver)
5. a valid texture handle               (live, never recycled)
       |
       v
   construct the transaction  ->  private receiver apply
```

Semantic eligibility deliberately sits **before** the structural walk. It is two
Automation property reads, about two microseconds, and it means a dangerous Shape
class is turned away before anything internal is touched at all.

## Proving the refusal is early enough

"PowerPoint did not crash" is not evidence that a refusal happened before the
dangerous call - a late refusal might survive by luck on one run.

So the backend counts entries into `ApplyCachedImage`, the function that builds
the OART transaction and calls the private receiver. `tools/test_semantic_guards.ps1`
reads the counter, attempts an apply, and reads it again:

```text
Case      Result Eligibility     ShapeType Entries Healthy
Connector Passed Unsupported     1         0 -> 0  yes
Line      Passed Unsupported     9         0 -> 0  yes
WordArt   Passed NativeSupported 1         0 -> 1  yes
```

`0 -> 0` is the proof: the private apply was never entered. `0 -> 1` for WordArt
shows the same test would notice if a refusal were missing. These cases are
permanent.

## Falling back is a property of the class

```text
NativeSupported    -> accelerated apply
FallbackSupported  -> Office's own Fill.UserPicture
Unsupported        -> BB_E_UNSUPPORTED_SHAPE
Invalid            -> BB_E_INVALID_SHAPE
```

The route is chosen from the verdict, never from whether something failed. There
is no "try native, and if it throws, try the slow way" anywhere in this codebase,
because that would hide an unvalidated Office build, a deleted Shape, a corrupt
image, an invalid handle or a bug in the backend behind a path that merely looks
slower. Those are exactly the failures worth surfacing.

`BB_E_INVALID_SHAPE` and `BB_E_UNSUPPORTED_SHAPE` are separate for the same
reason: one is a broken object, the other is a fact about PowerPoint.

## This is one Office build

Every semantic verdict here is a property of **16.0.14334.20848**. Nothing about
it is assumed to generalise. On any other build the structural guards refuse
first, so an unvalidated build never reaches these decisions - and enabling one
means repeating all three:

```text
structural checks  +  semantic class validation  +  the compatibility harness
```

`tools/test_shape_compatibility.ps1` runs every Shape category in its own
PowerPoint process, so a class that kills the host costs one row rather than the
run, and records `CrashedDuringResearch` for it. Production code never relies on
that category: a class known to be dangerous is converted into a clean refusal
long before any private call.
