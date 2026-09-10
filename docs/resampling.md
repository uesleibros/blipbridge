# Image resampling

BlipBridge can resample a raw BGRA32 image to any size before turning it into a
texture, with an explicit choice of filter.

```vb
Dim texture As BlipBridgeTexture
texture = BlipBridge.LoadTexturePixelsScaled( _
    pixels, sourceWidth, sourceHeight, sourceStride, _
    targetWidth, targetHeight, BBScaleBicubic)
```

or through the C ABI:

```c
BB_LoadTexturePixelsScaled(pixels, width, height, stride,
                           targetWidth, targetHeight, BB_SCALE_BICUBIC, &handle);
```

## What this controls, and what it does not

```text
source BGRA  ->  BlipBridge filter  ->  image resource  ->  Office
                 ^^^^^^^^^^^^^^^^^                          ^^^^^^
                 this is what you choose                    this is not
```

BlipBridge resamples the pixels **before** Office receives them. It does not
change how Office then draws the resulting image: Shape scaling, slideshow
scaling, zoom, DPI and the renderer's own sampling are all downstream and
unaffected by anything here.

The supported claim is exactly this: *BlipBridge provides explicit software
image-resampling filters before the image is handed to Office.*

**Office may still resample the image when it draws it.** If a Shape is not
exactly the pixel size of the image, or the slide is zoomed, or a slideshow is
scaled to the display, PowerPoint's own renderer scales it with its own sampling
- which BlipBridge neither chooses nor sees. Picking `BBScaleBicubic` here does
not make PowerPoint's final draw bicubic.

The practical consequence: to get exactly the pixels you chose, the image has to
reach the renderer at the size it will be drawn. That is up to the caller and the
document, not this library. Whether Office exposes any control over its own
sampling is a separate question that has not been investigated.

## The filters

Three, and every one is implemented and tested. A filter is not given a name
until it works, so there is nothing here to discover as a placeholder at run
time.

| Value | VBA | Method |
|---|---|---|
| `BB_SCALE_NEAREST` | `BBScaleNearest` | exact point sampling |
| `BB_SCALE_BILINEAR` | `BBScaleBilinear` | 2x2 linear interpolation |
| `BB_SCALE_BICUBIC` | `BBScaleBicubic` | Catmull-Rom cubic over 4x4 |

**Nearest** copies the chosen source pixel whole. No averaging, so the source
BGRA survives byte for byte, alpha included. Doubling `A B / C D` gives exactly:

```text
A A B B
A A B B
C C D D
C C D D
```

which the test suite asserts against those literal bytes rather than a tolerance.

**Bilinear** interpolates between the four source pixels surrounding the sample
point.

**Bicubic** uses **Catmull-Rom** - the `a = -0.5` member of the Keys family. It
is named exactly because "bicubic" describes a family, not a result: Catmull-Rom
interpolates, so the curve passes through the source samples and an unscaled
image is unchanged, where Mitchell-Netravali is smoother but does not
interpolate and B-spline blurs. It overshoots slightly at sharp edges, which is
inherent to the kernel and is clamped to the byte range rather than left to wrap.

## Conventions, applied identically to every filter

**Pixel centres.** A destination pixel `d` samples the source coordinate
`(d + 0.5) * scale - 0.5`. Both images are treated as grids of unit squares with
samples at their centres, which is what stops a scaled image drifting half a
pixel. Filters that each chose their own convention would produce images offset
from one another; this one is shared.

**Edges** are clamped to the nearest valid source pixel. No wrapping, no
mirroring, no invented border colour.

**Alpha** is interpolated in premultiplied form by the interpolating filters,
then converted back to straight BGRA. This is not a detail: interpolating
straight BGRA blends the colour of fully transparent pixels into visible ones,
which is precisely the dark halo that appears around cut-out edges. Weighting
colour by alpha first makes a transparent pixel contribute nothing but its zero
alpha. Nearest does no blending at all, so it is exact either way.

The test suite asserts this directly - a transparent black pixel beside an opaque
white one must not darken the white as it is upscaled.

**Same size in and out** copies the rows and runs no filter, so the result is
bit-exact whichever filter was named, and costs about as much as a memcpy.

## Validation

Every size and stride is checked, and every product is computed in 64 bits and
range-checked, before anything is read or allocated:

| Rejected | Why |
|---|---|
| null source | nothing to read |
| zero source or target width or height | no meaningful image |
| stride below `width * 4` | rows would overlap or read past the buffer |
| zero or negative stride | not a layout |
| unknown filter value | no placeholder filters exist |
| more than 64 megapixels either side | a limit is what stops a bad width and height becoming a wild allocation |

Padded strides are honoured, so a source whose rows are aligned works without
being repacked first. The test suite fills that padding with `0xCD` so a filter
that ever read past a row would produce a visibly wrong result rather than a
plausible one.

## Performance

Measured on the development machine, milliseconds per resample, scaling only -
no cached-image creation and no Office apply, which are timed separately.

| Case | nearest | bilinear | bicubic |
|---|---:|---:|---:|
| 32x32 → 128x128 | 0.014 | 0.338 | 0.911 |
| 128x128 → 32x32 | 0.001 | 0.021 | 0.064 |
| 256x256 → 512x512 | 0.243 | 5.89 | 15.40 |
| 512x512 → 256x256 | 0.059 | 1.39 | 3.64 |
| 512x512 → 1920x1080 | 2.09 | 47.5 | 186.1 |
| 1920x1080 → 512x512 | 0.49 | 10.3 | 24.3 |
| 100x75 → 333x251 | 0.126 | 3.11 | 7.70 |
| 512x512 → 512x512 (identity) | 0.066 | 0.053 | 0.067 |

Reproduce with `build/<config>/bb_resample_benchmark.exe`.

Read these plainly. **Nearest is roughly twenty times cheaper than bilinear and
eighty times cheaper than bicubic**, and bicubic upscaling to full HD costs
around 186 ms.

### When you pay it

**Scaling cost is paid on every `BB_LoadTexturePixelsScaled` call.** It is not
inherently a one-time cost, and whether it behaves like one depends entirely on
what the caller does:

* **Static or reused images.** Scale once, keep the handle, apply it as often as
  you like. The cost is then paid once and amortised over every apply - and at
  that point 186 ms for a full-HD bicubic upscale is usually irrelevant.
* **New content each time.** A workload that builds a *new* scaled texture per
  frame pays the full scaling cost per frame, on top of cached-image creation and
  the apply. At 186 ms, bicubic is not viable there; nearest, at 2 ms, may be.

Nothing caches by pixel content, so calling `LoadTexturePixelsScaled` twice with
the same buffer does the work twice. If you want it once, keep the handle.

### Which stage is which

The four costs are measured separately and should not be added up carelessly:

| Stage | Where it is measured |
|---|---|
| scaling | `bb_resample_benchmark`, the table above |
| cached-image creation | `docs/cost_profile.md`, the `create` stage |
| `ApplyTexture` | `docs/benchmarks.md` and `cost_profile.md`, the `apply` stage |
| whole `LoadTexturePixelsScaled` | scaling plus cached-image creation |

A single number covering all of them would hide which half a change affected.

The identity row confirms the fast path: all three filters cost the same there,
because none of them runs.

One optimisation is present and was worth measuring: the horizontal weights and
clamped source columns depend only on the destination column, so they are
computed once per column rather than once per pixel. That took bicubic full-HD
upscaling from 267 ms to 186 ms and bilinear from 66 ms to 48 ms. No SIMD and no
threading were added, because nothing has shown them to be necessary and both
would trade host stability for speed in a library that already runs inside
PowerPoint's process.

## Testing

`tests/resample_contract.cpp` runs in CI on both architectures with no Office. It
covers exact nearest output, 1x1 to arbitrary sizes, odd and non-square targets,
non-integer ratios, upscaling and downscaling, flat-field preservation, identity,
padded stride, transparent-edge handling for both interpolating filters,
source-sample preservation on an exact upscale - which is what pins the
pixel-centre convention - and every rejection above.
