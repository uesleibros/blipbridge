# Raw pixels, and whether images can be mutated

Two questions, both answered by measurement rather than by design intent:

1. Can a texture be built from raw pixels, skipping image encoding entirely?
   **Yes** - and it is documented below, including the part where it is *not*
   faster than it sounds.
2. Can an existing image resource have its pixels updated in place, so a
   changing picture avoids creating a new one every frame? **No**, on the
   evidence available.

## Raw pixels work

GFX exports a second cached-image creator taking a pixel buffer rather than a
stream, at `GFX +0x193C00`:

```c
TCntPtr<ICachedImage> Create(TCntPtr<IImage>& image, const void* pixels,
                             unsigned width, unsigned height, int stride,
                             ARC::SurfaceFormat format,
                             const Math::TVector2<float>& dpi);
```

Its prologue is the same shape as the stream creator: RCX is the hidden sret
pointer, the result is *moved* into it rather than AddRef'd - so the returned
pointer carries exactly one reference - and a temporary is released through
`GFX +0xA1A0` on the way out. Ownership is therefore identical to the encoded
path, and the apply path needs no changes at all: a cached image is a cached
image, wherever it came from.

### The surface format carries no information

`ARC::SurfaceFormat` has no symbols, so it was probed rather than guessed.
Feeding a 32bpp buffer of an asymmetric colour - R=F0, G=40, B=10, chosen so a
red/blue swap is unmissable - and creating with every value from 0 to 24:

**Every value rendered the same correct BGRA result.**

That is a claim about the argument mapping as much as about the format, so the
mapping was proved separately. A two-colour image renders as two halves only if
width and stride land where intended, and a deliberately wrong stride must skew
it:

```text
correct stride  : L=F04010 R=1040F0     (left red, right blue, as authored)
stride + 4      : L=1040F0 R=F04010     (halves swapped - stride is honoured)
```

So width, height and stride do reach their parameters, and the format value
genuinely is tolerated across that range for a 32bpp buffer.

Consequently the public API offers **BGRA32 only** and passes a fixed value. A
`format` parameter that accepts anything and means nothing would be worse than no
parameter: it would look like a contract. Other layouts are the caller's to
convert.

### It is not faster than loading an encoded image

Measured in process through the public ABI, 200 iterations, each texture
released immediately, comparing against a PNG of the **same dimensions**:

| Size | encoded load | raw pixel load | ratio |
|---|---:|---:|---:|
| 32x32 | 0.0151 ms | 0.0148 ms | 1.02x |
| 64x64 | 0.0240 ms | 0.0280 ms | 0.86x |
| 128x128 | 0.0642 ms | 0.0764 ms | 0.84x |
| 256x256 | 0.2315 ms | 0.2837 ms | 0.82x |

Raw pixels are consistently **15-18% slower**. Office's PNG decode is quick, and
the raw path still copies and converts the whole buffer, which scales with pixel
count: 640x480 costs 1.27 ms against 0.0148 ms for 32x32, roughly linear in
pixels.

So the raw path is not a speed feature. Its value is narrower and real: a caller
that already holds pixels does not have to **encode a PNG first**, which costs
far more than either row above. Choose it when you have pixels; choose
`BB_LoadTexture` when you have a file.

The 320x288 and 640x480 rows in the raw benchmark are marked `comparable=0`
because no PNG of those dimensions exists in the repository, so the encoded leg
there is decoding a smaller image and the ratio is meaningless. They are kept
only as raw-path scaling data.

## Mutation is not available

The ideal for changing content would be: create once, bind once, then update
pixels in place. That does not appear to be possible.

**No export mutates an image.** Enumerating all 880 GFX exports and searching for
update, lock, replace, set-pixel or invalidate operations on `IImage` or
`ICachedImage` finds nothing that writes to one. What exists is:

| Export | Direction |
|---|---|
| `ICachedImage::Create` (6 overloads) | create |
| `IImage::Create` | create |
| `Gfx::CreateImageFromHBITMAP` | create |
| `Gfx::CreateHBITMAPFromImage` | read out |
| `ICachedImage::GetUniqueID` | read |
| `ICachedImage::FCachingEnabled` | read |

`IImage::GpImageLock` exists as a nested type, but only its copy constructor and
destructor are exported - the constructor that would produce a lock from an image
is not, so there is no reachable way to obtain one.

**The design says immutable too.** `ICachedImage` carries a *unique ID* and a
caching flag. An object that is content-addressed and cached by identity cannot
safely have its content changed underneath: every cache keyed on that ID would
then be wrong. The absence of a mutator is consistent with the object being
deliberately immutable rather than merely lacking a convenience.

This is evidence, not proof. It rules out the exported surface; it does not rule
out an unexported path. But guessing at one would mean writing into a shared,
content-addressed cache, which is exactly the class of change this project
refuses to make without proof.

### What that leaves for changing content

Create a new texture per frame, apply, release.

> **Corrected.** This section first reported ~0.9-1.0 ms per frame, measured over
> 2,400 iterations. That measurement was taken **from PowerShell**, which pays a
> cross-process Automation round trip per call and made three of them per frame.
> Re-measured in process through the public C ABI, a new-content frame costs
> **0.31-0.34 ms**: 0.08 ms to create a 128x128 BGRA image, 0.21-0.24 ms to
> apply it, and 0.0007 ms to release it. The stage-by-stage breakdown is in
> [cost_profile.md](cost_profile.md).

Against 0.19 ms for applying an existing texture, changing content therefore
costs roughly 1.6x an ordinary apply, not 5x. It is still more, and the extra is
almost entirely the image creation - the apply itself barely notices whether the
image is new.

That makes the honest answer: **BlipBridge is excellent for a fixed set of images
reused many times, and usable for content that changes every frame** - a
new-content frame is still about twice as fast as `Fill.UserPicture` is for a
*repeated* one.

Ring buffers and pooling of cached images were measured rather than assumed. A
pool of pre-created images is safe - reference counts stay bounded and applying
from one costs exactly what any other apply costs - but without mutation it can
only recycle handles, never work, so it helps only when the set of distinct
images is bounded and recurring. See `cost_profile.md`; no pooling API was added,
because a caller with a bounded image set already expresses it by holding those
handles.

## No slowdown over time was reproduced

A "fast at first, gradually slower, eventually stabilises" symptom was reported
for this class of workload. Three configurations were run to find it, measuring
apply time in buckets while tracking reference counts, cached-image creations,
handle count and private bytes:

| Configuration | 2,400-4,000 applies | Drift |
|---|---|---|
| one texture, undo accumulating | 0.466 -> 0.438 ms | 0.87x - 1.00x, flat |
| one texture, `StartNewUndoEntry` per apply | 0.574 -> 0.753 ms | up to 1.31x, and slower overall |
| a new texture every apply | 0.929 -> 0.877 ms | 0.92x - 1.27x, noisy but flat |

The absolute numbers in that table are cross-process Automation timings, so they
are several times the in-process cost; what the table is evidence for is the
*drift* column, which is what it was built to measure.

**The symptom did not reproduce.** Reference counts stayed pinned at their
attributed values, creations tracked `LoadTexture` calls exactly, handles
returned to baseline, and private bytes grew ~20 MB across thousands of applies
and then fell back on their own.

One useful negative result: calling `StartNewUndoEntry` before each apply, an
obvious-looking way to stop undo history growing, makes things **slower** and
adds drift. It is an anti-fix.

So the cause of the reported symptom is not in any of these paths. Reproducing it
needs the workload that showed it - most likely something this test does not do,
such as growing slide content, many distinct Shapes, or document growth from
another source. Instrumenting that workload with the same four counters is the
next step; the counters are already exposed through `InspectTexture`.
