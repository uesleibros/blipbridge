# Which Shapes BlipBridge can fill

The backend used to accept two Shape classes, AutoShape and Freeform, because
those were the two whose receiver chain had been proved. That was honest but
narrow, and it was never established what the other classes actually do.

This is the answer, measured on **PowerPoint LTSC 2021 x64, build
16.0.14334.20848**. Every row was produced by creating a real instance of the
class, walking the fill chain, applying, and verifying - not by reasoning about
what ought to work.

## The matrix

Two columns matter and they are independent: **structural** is how far the Shape
gets along the validated internal chain, **semantic** is what the production
policy decides about its class. Connector and Line are `Complete` structurally
and `Unsupported` semantically, which is the whole point - see
[safety_model.md](safety_model.md).

| Category | Class | `Shape.Type` | Structural | Semantic |
|---|---|---|---|---|
| AutoShape | **NativeSupported** | 1 msoAutoShape | Complete | NativeSupported |
| Freeform | **NativeSupported** | 5 msoFreeform | Complete | NativeSupported |
| TextBox | **NativeSupported** | 17 msoTextBox | Complete | NativeSupported |
| Placeholder | **NativeSupported** | 14 msoPlaceholder | Complete | NativeSupported |
| Callout | **NativeSupported** | 2 msoCallout | Complete | NativeSupported |
| Group | **NativeSupported** | 6 msoGroup | Complete | NativeSupported |
| Group child | **NativeSupported** | 1 msoAutoShape | Complete | NativeSupported |
| Picture | **NativeSupported** | 13 msoPicture | Complete | NativeSupported |
| WordArt | **NativeSupported** | 1 msoAutoShape | Complete | NativeSupported |
| Media | **NativeSupported** | 16 msoMedia | Complete | NativeSupported |
| Table | FallbackSupported | 19 msoTable | NoReceiver | FallbackSupported |
| SmartArt | FallbackSupported | 24 msoIgraphic | NoReceiver | FallbackSupported |
| OLE object | FallbackSupported | 7 msoEmbeddedOLEObject | NoReceiver | FallbackSupported |
| Chart | Unsupported | 3 msoChart | NoReceiver | Unsupported |
| **Connector** | **Unsupported** | 1 msoAutoShape | **Complete** | **Unsupported** |
| **Line** | **Unsupported** | 9 msoLine | **Complete** | **Unsupported** |

<details><summary>The earlier one-column form</summary>

| Category | Class | `Shape.Type` | Native path |
|---|---|---|---|
| AutoShape | **NativeSupported** | 1 msoAutoShape | validated chain |
| Freeform | **NativeSupported** | 5 msoFreeform | validated chain |
| TextBox | **NativeSupported** | 17 msoTextBox | validated chain |
| Placeholder | **NativeSupported** | 14 msoPlaceholder | validated chain |
| Callout | **NativeSupported** | 2 msoCallout | validated chain |
| Group | **NativeSupported** | 6 msoGroup | validated chain |
| Group child | **NativeSupported** | 1 msoAutoShape | validated chain |
| Picture | **NativeSupported** | 13 msoPicture | validated chain |
| WordArt | **NativeSupported** | 1 msoAutoShape | validated chain |
| Media | **NativeSupported** | 16 msoMedia | validated chain |
| Table | FallbackSupported | 19 msoTable | no receiver chain |
| SmartArt | FallbackSupported | 24 msoIgraphic | no receiver chain |
| OLE object | FallbackSupported | 7 msoEmbeddedOLEObject | no receiver chain |
| Chart | Unsupported | 3 msoChart | no receiver chain, and `UserPicture` also refuses |
| Connector | **Unsupported - refused by name** | 1 msoAutoShape | chain present, but fatal |
| Line | **Unsupported - refused by name** | 9 msoLine | chain present, but fatal |

</details>

`FallbackSupported` means the native path is refused but ordinary
`Fill.UserPicture` works, so the `UserPicture2` dispatcher has somewhere to go.
`Unsupported` means neither path produces a picture fill. The research harness
also has `NotApplicable`, for a class that could not be created on the machine,
and `CrashedDuringResearch`, for one whose host did not survive the test -
kept as historical evidence and never relied on by production code, which turns a
known-dangerous class into a clean refusal before any private call.

Reproduce with `tools/test_shape_compatibility.ps1`; the machine-readable form is
`artifacts/shape_matrix.txt`.

## The connector, and why structure is not enough

This is the single most important result here, and it was a surprise.

A **Connector reports `Shape.Type = 1`, msoAutoShape.** It presents the same
PPCORE delegating wrapper (`ppcore.dll+0x1464478`, inner offset `0x8`, 31 identity
thunks), the same inner OART FillFormat vtable (`oart.dll+0xAF60B8`), the same
control block and the same receiver vtable as an ordinary rectangle. Every
structural check in `ResolveFillTarget` passes. The read-only classifier reports
`step=Complete`.

**Applying a texture to one terminates PowerPoint.**

Isolated in `tools/test_connector_isolation.ps1`, three cases in three fresh
processes:

| | result |
|---|---|
| AutoShape, native apply | survives, `Fill.Type = 6` |
| Connector, `Fill.UserPicture` | **Office itself refuses**: "value out of range" |
| Connector, native apply | **RPC 0x800706BE - the host is gone** |

The middle row is the explanation. Office performs a validity check in PPCORE
*before* the fill handler is ever reached, and a connector fails it: a connector
is a line, and a line has no interior to fill. The native apply reproduces the
handler faithfully and does not reproduce that pre-check, so it drives the
handler into a state Office never allows it to reach.

Two consequences, both now built in:

1. **`Shape.Type` alone is not a safety check**, because classes lie about it.
   Connector and WordArt both report msoAutoShape and behave completely
   differently.
2. **The structural walk is necessary but not sufficient.** Passing it means the
   objects are the ones we think they are. It says nothing about whether the
   operation is meaningful for that Shape.

The guard is `Shape.Connector`, which is msoTrue for exactly Connector and Line
and msoFalse for every other class tested. It is Office's own answer to the
question, it costs about a microsecond, and a Shape that will not answer it is
refused rather than assumed safe.

This was a live defect: the shipping backend accepted `Shape.Type = 1` and would
have crashed PowerPoint for any caller who passed a connector.

## How a class earns native support

Structural evidence alone is not enough - the connector settled that. A class is
promoted only after all of this:

1. The read-only classifier reports the full validated chain
   (`tools/test_shape_compatibility.ps1`, which never calls a private function).
2. A native apply is attempted **in a disposable PowerPoint process**, through a
   research-only path that bypasses the type allowlist and nothing else.
3. The apply leaves identity, `Shape.Type`, geometry, rotation and Z-order
   untouched, produces `Fill.Type = 6`, and creates no Picture Shape.
4. The class then survives `tools/test_shape_class_stress.ps1`:

   * 1000 repeated applies of one texture
   * 500 alternating applies of two textures
   * undo and redo
   * save, close and reopen
   * deleting the Shape and applying to a fresh one
   * closing the presentation

   with the cached image's reference count returning to **1** - the handle's own
   - no leaked handles, and the host still able to run an ordinary
   `Fill.UserPicture` afterwards.

All ten native classes cleared every step. Reference counts after close were 1 in
every case, handles at end 0 in every case, host healthy in every case.

## The policy lives in one place

`ClassifyShapeForNativePictureFill` in
`src/backend/windows_office/shape_policy.cpp` is the only definition of what is
acceptable. The C ABI backend, the COM compatibility surface and the
`UserPicture2` dispatcher all ask it; none re-derives the answer, because both
reach the same texture store and it would be a defect for them to disagree.

It returns one of four verdicts - `NativeSupported`, `FallbackSupported`,
`Unsupported`, `Invalid` - from an allowlist of `msoShapeType` values plus the
connector refusal, which is checked *first* because a Connector reports
msoAutoShape. Adding a class means running the two harnesses above, not editing
the list.

The refusal is proved to happen before any private call:
`tools/test_semantic_guards.ps1` reads the backend's count of entries into the
OART apply, attempts one, and reads it again. Connector and Line report `0 -> 0`.
Those cases are permanent.

## What is not claimed

* This is **one Office build**. The matrix is a property of 16.0.14334.20848 and
  has to be re-run on any other.
* `NativeSupported` means the fill is applied safely and survives everything
  above. It does not promise the result is *visually useful* for every class - a
  picture fill behind a Media object's poster frame is legal, applied and
  persisted, but whether you want it is your call.
* Table, SmartArt and OLE have no validated chain at all. They are not "not yet
  supported pending work"; nothing in the current model reaches them.
* Chart is genuinely unfillable by either path on this build.
