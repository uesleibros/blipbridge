# Preparing an image before Office sees it

Research notes for the v0.7.0 image work. Everything here is measured on Office
**16.0.14334.20848** x64. No public API is frozen yet; this is the evidence the
API should be designed from.

## 1. What a texture actually is

Settled first, because three proposed features depended on the answer.

A `BlipBridgeTexture` holds **two opaque GFX pointers** - the cached image and
its companion image object - a byte count, an apply count, and a process-unique
id. It retains **no CPU pixels**. There is no path from a handle back to pixels
that does not mean reverse-engineering GFX's internal storage.

So every image operation is an **at-load** operation, from encoded bytes, a file,
or raw BGRA. There is no `CreateFlippedTexture(existingHandle)`, and there will
not be one until something retains source pixels. Promising handle-to-handle
processing would mean promising to read data Office does not expose.

## 2. Decoding

`src/image/decode.cpp` decodes through Windows Imaging Component to one canonical
format: **32bpp BGRA, straight (non-premultiplied) alpha, top-down, tightly
packed**. That is what the resampler documents as its input and what the
raw-pixel entry points already take, so a decoded image and a caller-supplied
buffer are the same thing downstream.

Straight rather than premultiplied alpha is deliberate: `resample.hpp`
premultiplies internally where it needs to, and doing it twice darkens every
transparent edge.

WIC converts from any source format in one step, so nothing here carries a switch
over pixel layouts. Formats **tested**: PNG with and without alpha, JPEG, BMP.
WIC handles more and they will very probably work, but an untested format is not
a supported one.

## 3. The stages, and their order

Crop, orient and resize are one journey, not three features, so they are one
request:

```
decode  ->  crop  ->  transform  ->  resize
```

Crop first, because a region is expressed in the source image's own coordinates
and that is the only frame a caller can name it in. Transform before resize, so
the target size always means the size of what comes out - rotating after resizing
would make the target describe an image the caller never sees.

Each stage switches itself off when unset, so the common case stays short and one
decode serves every stage. That matters more than it sounds: see the timings.

Transforms are the lossless ones only - flip horizontal, flip vertical, rotate
90/180/270. Every output pixel is exactly some input pixel, so pixel art survives
them. Arbitrary-angle rotation is absent on purpose: it needs resampling, a
background colour and an output-size decision, none of which are free choices.

## 4. How PowerPoint maps a picture fill

The measurement everything else rests on. `tools/probe_freeform_uv_mapping.ps1`
fills seven Freeforms with a UV gradient - red carries u, green carries v - and
reads back which source texel arrived at each of 81 interior points.

Against the hypothesis that the image is mapped linearly onto the Shape's
**bounding box** and then clipped to the path:

| quad | mean error | worst error |
|---|---:|---:|
| rectangle | 0.4 px | 0.7 px |
| wide rectangle | 0.4 | 0.8 |
| parallelogram | 0.4 | 0.6 |
| trapezoid | 0.4 | 0.7 |
| extreme trapezoid | 0.4 | 0.7 |
| rotated quad | 0.5 | 0.7 |
| perspective quad | 0.4 | 0.7 |

In source-texel units on a 256x256 probe, which is the quantisation floor of a
gradient with 256 levels. **Path geometry does not move the texture.** It decides
what is visible, not what is where.

That is the good outcome, and not the one to assume. Because Office's mapping is
known exactly and is simple, a caller-supplied quad can be compensated for.

## 5. Projective warp

`src/image/warp.cpp` maps the unit square onto four caller-supplied corners with
a real homography and rasterises into the quad's bounding box, leaving the
outside transparent. Warp into that box and Office's own mapping is the identity;
the path clips the rest.

Point order is fixed and never reordered behind the caller:

```
p[0] = source top-left      (u,v) = (0,0)
p[1] = source top-right             (1,0)
p[2] = source bottom-right          (1,1)
p[3] = source bottom-left           (0,1)
```

A caller who hands them in another order is asking for a mirrored or crossed
mapping and gets one - which is the only behaviour that lets someone mirror a
face on purpose.

**Verified in isolation** (`tests/warp_contract.cpp`): a rectangle warps to the
exact identity; corners land on corners; a parallelogram keeps the source centre
at the centroid, so the projective terms vanish on their own; an extreme
trapezoid puts the source mid-row well above half height with v monotonic;
outside the quad is transparent; duplicated, collinear, bow-tie, non-finite and
oversized quads are each refused by their own name before anything is allocated.

**Verified in PowerPoint** (`tools/test_quad_warp.ps1`): the same UV probe warped
onto a tapered Freeform and read out of the render. The source left and right
edges arrive at the quad's top nodes, where an unwarped box mapping would put u
at 0.43 and 0.71. The source mid-row lands at **22%** of the height, not 50%.
Nothing outside the quad is painted. The fill survives save and reopen.

Nothing reads `Shape.Nodes`. The caller already knows the points.

### What is compensated, and what is not

Compensated: where each source texel lands inside the quad. That is entirely our
arithmetic, done before Office sees the image.

**Not** compensated, and not ours: how Office then draws that image when the
slide is zoomed, scaled, printed or shown. Office's own sampler runs downstream
of everything here. The supported claim is about the image Office receives.

### Bicubic is not offered for the warp

In a projective sampler it costs sixteen taps per output pixel with a varying
kernel footprint. Nearest and bilinear are implemented; bicubic is refused by
name rather than silently downgraded, because offering it for API symmetry would
hide a cost nobody asked for in a per-frame path.

## 6. Dynamic textures: not viable by the cheap path

The question was whether the pixels behind an image a Shape already shows can be
changed without a new fill.

`GEL::ICachedImage::Create` **copies** the pixel buffer. The probe
(`experiments/exp_internal_blip/dynamic_probe.cpp`) keeps the source buffer
alive, applies it, overwrites it red-to-blue, and exports the Shape: it renders
red. Re-applying the same cached image - which distinguishes a stale repaint from
a copy - also renders red.

This does not prove no mutable path exists anywhere inside GFX. It proves the
obvious one does not, which is what decides whether to spend weeks looking for a
harder one. Recorded so the question is not re-asked from scratch.

## 7. Timings

One machine, one session, medians of repeated runs. Absolute times move with
machine load; the shape is the point.

| | 256x256 | 1024x1024 |
|---|---:|---:|
| decode PNG | 0.30 ms | 3.18 ms |
| decode JPEG | 0.51 | 7.79 |
| + half-size nearest | +0.07 | +2.19 |
| + half-size bilinear | +0.40 | +6.38 |
| + half-size bicubic | +0.92 | +14.74 |

| one more 16x16 tile to 256x256 from an already-decoded atlas | 0.06-0.07 ms |
|---|---:|

| warp onto a tapered quad | nearest | bilinear |
|---|---:|---:|
| 256x256 | 0.21 ms | 1.16 ms |
| 512x512 | 1.14 | 5.21 |
| 1024x1024 | 4.24 | 19.81 |

**Decoding dominates everything.** That is why the pipeline decodes once and runs
every stage from that result: a second tile out of the same atlas costs 0.06 ms
rather than another 3.18.

The warp fits a frame at useful sizes - 256x256 bilinear is 1.16 ms of a 16.67 ms
budget - and does not at 1024x1024 bilinear, which is 19.8 ms. "Fits" there means
the warp alone; a real frame also pays to create the image resource and for
Office's own apply.

## 8. x86

The image code is ordinary portable work with no Office in it, so the decode,
pipeline and warp suites build and run on x86 as well as x64. That is a statement
about **image processing**, and it does not move the existing line: the x86
Office backend is still not validated in a real 32-bit PowerPoint, and building
is not running. See `windows_x86.md`.
