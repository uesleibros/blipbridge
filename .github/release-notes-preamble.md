## Which package do I want?

Check **PowerPoint → File → Account → About PowerPoint**. The first line ends with
`64-bit` or `32-bit`, and that is the package to download. If you pick the wrong
one, `BlipBridge.bas` says so in those words rather than letting you meet a
loader error.

| Package | Native backend |
|---|---|
| `windows-x64` | **Validated** against Office 16.0.14334.20848 |
| `windows-x86` | **Not yet** — the library loads and refuses every texture call with a specific reason |

The x86 package is published so the ABI, the VBA wrapper and the packaging can be
exercised on 32-bit Office. It does not accelerate anything yet, and it is built
without the 64-bit backend compiled in, so it cannot. Progress:
[docs/windows_x86.md](https://github.com/uesleibros/blipbridge/blob/main/docs/windows_x86.md).

## Install

1. Copy `BlipBridge.dll` next to your `.pptm`.
2. Import `BlipBridge.bas` into the VBA project.

No `regsvr32`, no ProgID, no add-in.

```vb
BlipBridge.UserPicture2 shp, "C:\textures\brick.png"
```

## Verify a download

```
sha256sum -c SHA256SUMS.txt
```

## What is validated, and what CI proves

The accelerated backend depends on internal Office layouts checked per build. It
runs against **one** validated build and refuses every other one rather than
guessing. See
[docs/safety_model.md](https://github.com/uesleibros/blipbridge/blob/main/docs/safety_model.md)
for the structural/semantic distinction and
[docs/shape_compatibility.md](https://github.com/uesleibros/blipbridge/blob/main/docs/shape_compatibility.md)
for which Shape classes work.

CI proves the library **builds** on both architectures. It does **not** prove the
PowerPoint backend works — hosted runners have no Office. Those suites run on a
machine with the validated build, and their transcripts are committed under
`docs/evidence/`.
