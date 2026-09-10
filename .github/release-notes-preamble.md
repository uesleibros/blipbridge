## One package, both architectures

There is a single download. It contains both DLLs, and `BlipBridge.bas` loads the
one matching the **PowerPoint process** at run time — so you do not need to know
which Office you have before downloading, and both can sit beside the same
presentation.

| | Native accelerated backend |
|---|---|
| **x64** | **Validated** against Office 16.0.14334.20848 |
| **x86** | **Not validated.** The DLL builds and its tests pass, and the wrapper selects it correctly — but it has never run inside a real 32-bit PowerPoint. It contains no accelerated backend at all and refuses every texture call with a specific reason. |

CI proves the library **builds** on both architectures. It does not prove the
PowerPoint backend works: hosted runners have no Office. See
[docs/windows_x86.md](https://github.com/uesleibros/blipbridge/blob/main/docs/windows_x86.md).

## Install

1. Copy `BlipBridge-x64.dll` and `BlipBridge-x86.dll` next to your `.pptm` — or
   just the one matching your Office, if you prefer.
2. Import `BlipBridge.bas` into the VBA project.

No `regsvr32`, no ProgID, no add-in.

```vb
BlipBridge.UserPicture2 shp, "C:\textures\brick.png"
```

## Verify a download

```
sha256sum -c SHA256SUMS.txt
```

## What is validated

The accelerated backend depends on internal Office layouts checked per build. It
runs against **one** validated build and refuses every other one rather than
guessing. See
[docs/safety_model.md](https://github.com/uesleibros/blipbridge/blob/main/docs/safety_model.md)
for the structural/semantic distinction and
[docs/shape_compatibility.md](https://github.com/uesleibros/blipbridge/blob/main/docs/shape_compatibility.md)
for which Shape classes work.

Office integration suites run on a machine with the validated build; their
transcripts are committed under `docs/evidence/`.

## Resampling, and what it does not control

`BB_LoadTexturePixelsScaled` resamples raw BGRA32 with an explicit filter —
nearest, bilinear, or Catmull-Rom bicubic — **before** Office receives the image.

It does not change how Office draws it. If a Shape is not exactly the pixel size
of the image, or the slide is zoomed, PowerPoint's renderer resamples again with
its own sampling, which BlipBridge neither chooses nor sees.
