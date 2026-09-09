# Where the time actually goes

Applying an existing texture was measured at ~0.19 ms; creating new content and
applying it had been reported at ~0.9-1.0 ms. The obvious reading is that
creating an image costs ~0.7 ms and that closing that gap is the interesting
problem. Both halves of that reading turned out to be wrong, which is why this
was profiled before anything was optimised.

The profiler is `experiments/exp_internal_blip/stage_profiler.cpp`, driven by
`tools/run_stage_profiler.ps1`. It times every step of the apply through
`ApplyCachedImage`'s optional stage hook, which no production call installs.

## The correction: new content costs ~0.31 ms, not ~0.9-1.0 ms

The 0.9-1.0 ms figure came from `tools/test_slowdown_attribution.ps1`, which
drives the work **from PowerShell**. That costs a cross-process Automation round
trip per call, and the changing-content leg made three of them per frame -
`LoadTexture` (marshalling a `Byte[]` as a `SAFEARRAY`), `ApplyTexture`,
`ReleaseTexture`.

Running the identical sequence in process through the real public C ABI:

| | per frame |
|---|---:|
| `BB_LoadTexturePixels` (128x128 BGRA, new content every frame) | 0.080 ms |
| `BB_ApplyTexture` | 0.21-0.24 ms |
| `BB_ReleaseTexture` | 0.0007 ms |
| **total** | **0.31-0.34 ms** |

So most of the reported gap was the measuring harness, not Office. A VBA caller
goes through a `Declare` transition rather than cross-process Automation, so it
pays close to the in-process number. The earlier figure is corrected in
`pixel_textures.md` rather than left standing.

## The stage breakdown

Medians, 400 iterations per leg, one Shape, 128x128 frames, taken with the
guard optimisation below in place:

| Stage | Cost | Share of an apply | Whose code |
|---|---:|---:|---|
| `resolve` - module guards, wrapper check, receiver walk | 0.022 ms | ~13% | ours |
| `record` - property record constructor, fill-kind slot | 0.0007 ms | <1% | Office |
| `subrecord` - image sub-record constructor | 0.0002 ms | <1% | Office |
| `install` - cached image into the sub-record | 0.003 ms | ~2% | Office |
| `transfer` - sub-record into the record's image slot | 0.001 ms | <1% | Office |
| `holder` - stretch holder and trailing flag | 0.0003 ms | <1% | Office |
| `transaction` - transaction constructor | 0.003 ms | ~2% | Office |
| **`apply` - the receiver call** | **0.15 ms** | **~80%** | Office |
| `released` - the four destructors | 0.001 ms | <1% | Office |

Creating new content adds `create` (0.08 ms for a 128x128 BGRA frame) and
`release` (0.0002 ms) on top.

### What that rules out

The ideas worth trying before this measurement existed were: reuse the record
allocation, reuse the image sub-record storage, pool the holder, avoid
constructor and destructor work.

**All of that together is under 4% of an apply - about 8 microseconds.** Record
construction, sub-record construction, both AddRef steps, the holder, the
transaction constructor and all four destructors sum to less than the noise on a
single `apply`. A perfect implementation of every one of those ideas would not
be measurable. They are not worth doing, and that is a measurement rather than an
opinion.

### What dominates

**About 80% of an apply is inside one call**: the receiver's transaction-apply
slot. That is Office mutating the document, writing the undo entry and marking
the Shape dirty. It is one private call; it is not decomposable from outside
without substantially more reverse engineering, and it is not something this
project should be trying to make cheaper - the undo entry and the invalidation
are the behaviour that makes a native apply indistinguishable from `UserPicture`.

**Repainting is not in these numbers.** The profiler loop never pumps messages,
so deferred rendering happens after the measurement. A workload that yields
between frames pays it; this measurement does not show it.

## The one optimisation the profile justified

`resolve` was 0.042 ms - over 20% of an apply - and it is entirely our code. Its
primitives, timed individually:

| Primitive | Cost | Calls per apply, before |
|---|---:|---:|
| `GetModuleInformation` | 0.0100 ms | 1 |
| `RequireSupportedModule` | 0.0095 ms | 2 |
| `GetModuleHandleW` | 0.0025 ms | ~10 |
| `DescribeDelegatingWrapper` (cached) | 0.0030 ms | 1 |
| `IsReadable` (one `VirtualQuery`) | 0.0003-0.0006 ms | 4 |
| `ResolveApplyFunctions` (cached) | 0.0001 ms | 1 |
| `Shape.Type` via `IDispatch` | 0.0010 ms | 1 |
| `Shape.Fill` via `IDispatch` | 0.0009 ms | 1 |

Two findings there. The Automation property fetches are **cheap** - about a
microsecond each, so `GetIDsOfNames` per call is not worth avoiding, which was
worth checking before assuming otherwise. And roughly 34 of the 37 microseconds
were Win32 module bookkeeping rather than any safety check: every
`RequireSupportedModule` call re-fetched all three module handles for the
cache's reload detector, and `GetModuleInformation` re-read PPCORE's
`SizeOfImage` on every single apply.

The fix keeps every check exactly as strict:

* `SynchroniseOfficeModules()` fetches the three handles **once per operation**
  and points the validation cache at them. The handles are the reload detector,
  so they still cannot be cached - but fetching them once instead of six times
  detects precisely the same thing.
* `SizeOfImage` is now cached per module handle, alongside the version check and
  the wrapper shapes that were already cached. It is a property of the loaded
  image, and the cache is discarded whenever a handle changes.

Nothing was removed, relaxed or made optional. The receiver is still re-resolved
on every apply, every vtable is still checked, every signature still verified.

**Result: `resolve` fell from 0.042 ms to 0.022 ms**, and the hot-path apply
median from 0.176 ms to 0.167 ms in `tools/run_texture_benchmark.ps1`. The
remaining ~0.022 ms is three `GetModuleHandleW` calls that cannot be cached
without losing reload detection, the wrapper lookup, and four `VirtualQuery`
readability probes.

## Pooling pre-created images

The question was whether a small pool of pre-created reusable image resources can
serve frequently changing content without unbounded growth. Measured with a ring
of eight distinct 128x128 images applied in turn, 400 applies:

| | |
|---|---|
| apply cost from the pool | 0.196-0.213 ms, the same as reusing one texture |
| reference counts after 400 applies | `4,5,6,7,7,7,7,8` - bounded, not climbing |
| private bytes across the run | +2 to +3 MB |

So a pool is **safe** - nothing accumulates per apply, and applying from a pooled
image costs exactly what applying any existing texture costs.

It is also **not a speed-up for changing content**, for the reason established in
`pixel_textures.md`: images cannot be mutated, so a pool can only recycle handles,
never work. A pool helps exactly when the set of distinct images is bounded and
recurring - a sprite sheet, a palette of states, an animation that loops. For
content that is genuinely new every frame, each frame still costs its own
`create`, and the pool contributes nothing.

That is why no public pooling or dynamic-texture API has been added. A caller who
has a bounded set of images already gets the whole benefit by holding those
texture handles, which the existing API does. Adding a pool type would be an
abstraction over something callers can already express.

## Reproducing

```powershell
.\tools\run_stage_profiler.ps1 400
```

Output goes to `artifacts/stage_profile.txt`, one `key=value` per line. Run it
more than once: single runs vary by 10-20% with machine load, and the medians are
what to read, not the means.
