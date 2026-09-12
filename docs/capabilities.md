# What the capability flags mean, and the evidence behind each

`GetCapabilities` is a statement about **the process it is called in**, not a
compile-time promise. The native backend is guarded to one exact Office build and
fails closed everywhere else, so a fixed string would be wrong on most machines.

Measured on Office 16.0.14334.20848 x64:

```text
outside PowerPoint : MemoryImageToFill=False;CachedTextureApply=False;PickUpFallback=True;FillOnly=False;InternalBackend=False
inside  PowerPoint : MemoryImageToFill=True; CachedTextureApply=True; PickUpFallback=True;FillOnly=False;InternalBackend=True
```

The three native flags move together, because they are all conditioned on the
same probe: PowerPoint host, `oart.dll`/`ppcore.dll`/`gfx.dll` all at the
validated build, the exported GFX creator present, and every private entry
point's signature bytes intact. Any failure means `False`.

## Two backends, and the one bit that separates them

There are two Office backends, and the capability mask is how a caller tells them
apart:

| | mask | `BB_CAP_NATIVE_BACKEND` |
|---|---|---|
| accelerated (x64, validated Office build) | `0x03FF` | set |
| portable (x86, or a forced-portable build) | `0x03FE` | clear |

**Every other bit is identical, and that is deliberate.** A capability bit says
whether a feature *works*, not whether it is fast. The portable backend really
does load textures from bytes, raw pixels and files, really does run the whole
image pipeline including the quad warp, really does apply to a Shape, a batch, a
range and conditionally with a skip cache, and really does serve
`BB_ApplyPicture`. So those bits stay set. Clearing them to signal "this is the
slow one" would tell callers a feature is missing when it is present, and the
feature-detection story would stop meaning anything.

`BB_CAP_NATIVE_BACKEND` is the bit that means "the accelerated Office-native
implementation is active". It is the only one that differs, and it is not a
synonym for "BlipBridge works" - see the `IsAvailable` against `IsAccelerated`
distinction in the README. A caller that tests the native bit to decide whether
to use BlipBridge at all would refuse a 32-bit install that works.

Both masks are asserted from inside PowerPoint by
`tools/test_portable_matrix.ps1`, which reads which backend answered and requires
the matching mask. Outside PowerPoint the mask is zero on both, because no
capability here is true without a host.

## MemoryImageToFill

**Claim.** Bytes already in memory reach a Shape fill with no temporary image
file.

**Why it was False before.** Milestone 4 measured the old `MemoryFillExperiment`
routing through `Fill.UserPicture` and found Office writing and re-reading a PNG
under `INetCache/Content.MSO`. Delivering the bytes was not the same as avoiding
the file.

**Evidence now.** `tools/test_no_temp_image.ps1` watches the Office image cache
with a `FileSystemWatcher` - a before/after listing misses files Office creates
and deletes between snapshots, which is why two earlier attempts read zero for
everything. With a working control:

| Leg | Content.MSO events |
|---|---:|
| 12x `Fill.UserPicture`, fresh source path each time | **24** |
| 12x native `ApplyTexture` | **0** |

The control had to use a *fresh path per call*, because Office caches by source
path and a repeated path hits the existing entry silently. Structurally this is
what the disassembly predicts: the native path never calls the file loader
`OART +0x950070`, it hands an `IStream` straight to the exported cached-image
creator.

## CachedTextureApply

**Claim.** One decoded image is applied to many Shapes without re-decoding.

**Evidence.** The store counts decodes. Across the stress run, `creations`
stayed exactly equal to the number of `LoadTexture` calls while applies ran into
the tens of thousands: 10,000 applies on one Shape and 2,000 alternating applies
added **zero** creations. 120 Shapes over 3 slides were filled from a single
texture. See `native_texture.md`.

## InternalBackend

**Claim.** The backend uses validated Office-internal entry points.

**Evidence.** Thirteen private OART functions plus one GFX export, each guarded
by module version, structural PPCORE wrapper validation, OART/GFX vtable checks
and per-function signature bytes. The ABI table with its evidence is in
`oart_abi.md`.

Note this flag does **not** mean the internal backend has replaced the fallback.
`GetBackendName` still reports `PickupApplyFallback`, and both paths are live at
once: `ApplyTexture` takes the native path for a native handle and the donor path
for a donor handle, distinguished by a disjoint handle range.

## PickUpFallback / FillOnly

`PickUpFallback=True` is unchanged: the donor path still exists and still works.

`FillOnly=False` stays false, and it is the one flag the native backend does
**not** satisfy. The donor path transfers whole-shape style; the native path
changes only the fill, so in the native sense the claim is arguably true - but
the flag has historically described the *fallback* backend, where it is false,
and one string cannot mean two things for two backends at once. Rather than
silently redefine it, it stays false; a per-backend capability string is the
right fix if this ever matters to a consumer.

## What is still not claimed

* **A build-independent backend.** Everything is pinned to one Office build, by
  design. On any other build every native flag reads `False` and the fallback is
  what remains.
* **Thread affinity beyond STA.** The OART FillFormat's reference count is a
  non-atomic increment; the store refuses calls from any thread but the one that
  created the first texture.
