# Examples

| File | Shows |
|---|---|
| `BasicUsage.bas` | loading bytes, one decode across many Shapes, an alternating two-frame loop, and the batch form |

To run one:

1. Copy `BlipBridge.dll` next to your `.pptm`.
2. Import `vba/BlipBridge.bas` and `examples/BasicUsage.bas`.
3. Put `texture.png` (and `frame1.png` / `frame2.png` for the animation) beside
   the presentation.
4. Run `Demo`.

Every example releases its textures in a `Cleanup` label, including on error.
That matters: a handle holds a decoded image until it is released or until
`Shutdown` runs.
