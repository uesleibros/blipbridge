# GFX resource retention in OART

Observed on Office 16.0.14334.20848 x64, 2026-09-09. This is read-only debugger
research, not a callable internal backend. Neither a resource pointer nor an
internal reference is retained by BlipBridge after the debugger detaches.

## Experiment and safeguards

`tools/prepare_decoder_trace.ps1` creates a disposable normal AutoShape, warms
decoder modules, records loaded versions, and waits for the debugger trigger.
The unique UserPicture operation then runs, followed by SaveAs and Close. A
separate text marker identifies those phases; no timings from this run are used.

Run GDB with `-ex "python probe_mode='lifetime'" -x tools/probe_decoder_stream.py`.
`consumer` mode stops after the OART consumer; `lifetime` follows the returned
resources until a tracked Release enters with count 1. The observer stops before
destruction, removes its breakpoints, and detaches. It never follows freed memory
or calls AddRef/Release itself. GFX/OART versions and breakpoint bytes are checked.
The first twelve cached-object AddRef callers are recorded to bound stack output.

## Confirmed observations

The cached result returned by GFX +0x7680 has vtable GFX +0x409DC0. The separate
image output has vtable +0x4055C8. Both use intrusive count at object +8, increment
through vtable slot 0 and decrement through slot +8. This is not IUnknown.

| Checkpoint | Cached count | Image count |
|---|---:|---:|
| Factory return | 1 | 2 |
| OART consumer entry | 1 | 2 |
| After consumer and local cached release | 1 | 2 |

OART +0x8F94C receives a record in RCX and a cached smart-pointer reference in RDX.
Its record member +0xF0 starts null and contains the exact cached pointer afterward.
The record's first qword is zero in this run; it must not be described as an object
with a known vtable. The observed address is near the active call's stack storage.
The consumer also obtains image metadata, so it is not a simple Shape fill setter.

The consumer's reference increment returns through OART +0x8F944, then +0x8FA95.
The caller releases its local cached reference at return address +0x8F575 with
count-before 2. The remaining reference is carried onward through OART state.

Further cached AddRefs return through OART +0x9848C. Disassembly at +0x9846D shows
source record +0xF0 copied to destination record +0xF0, followed by slot-0 increment.
This is in the record-copy routine entered at +0x97E50. Observed caller chains include:

```text
OART +0x89CA06 -> +0x22BD06 -> +0xA0898 -> +0x9848C -> GFX +0x90AA0
OART +0x89CA4D -> +0x488E8 -> +0x4893F -> +0x13B6E -> +0xA0898 -> +0x9848C
```

These are return-address RVAs. +0x89CA06 immediately follows the loaded-record
consumer call at +0x89CA01, after filename loader +0x950070 returns to +0x89C9F6.
This connects GFX retention with the existing UserPicture transaction, but does
not identify a standalone assignment ABI or distinguish document state from undo
copies for every reference.

With phase markers enabled, the cached object remains alive through UserPicture
and SaveAs. Its final observed Release enters GFX +0xA1A0 with count-before 1 during
`Presentation.Close`, from OART +0x3DBB1. This corrects the tentative concern that
the returned cached object might be discarded at the end of the loader call.
It is retained beyond that call; independent application-owned lifetime remains
untested. The image object's final destruction is not observed because the probe
stops before the cached object's last decrement.

## Persistence and limits

The ordinary traced UserPicture deck saved and reopened successfully: one normal
AutoShape, one picture fill, zero Picture shapes, one media PNG. This validates
the existing Office path while observed. It does not validate a synthesized GFX
resource, file-free loading, direct fill binding, or resource sharing between
multiple Shapes. Office still uses its Content.MSO temporary image in this path.

Next: observe the loaded-record transfer at OART +0x22BCF4 and the subsequent
transaction call that changes Shape state. Determine which non-image members are
needed for document ownership/serialization before attempting to supply our own
stream-created resource. A cached GFX pointer alone is not established as sufficient.

Evidence: `docs/evidence/oart_consumer.txt`, `oart_lifetime_phases.txt`,
`oart_retention.txt`, `decoder_lifetime_reopen.txt`, and
`docs/decoder_lifetime_package.md`. Nearest-export names in raw GDB stacks are
not reliable function names; use the RVAs above.
