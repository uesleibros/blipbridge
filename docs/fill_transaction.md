# Loaded image to OART transaction

Observed on OART 16.0.14334.20848 x64, 2026-09-09. The experiment is
`experiments/exp_internal_blip/probe_fill_transaction.py`. It sets validated
breakpoints during the existing UserPicture call; it does not invoke internal
methods or retain pointers after detach.

## Resource propagation

The same cached-image address was observed in each of these live records:

| Stage | Offset | Evidence |
|---|---:|---|
| Loaded image record | +0xF0 | Input RDX at OART +0x22BCF4 |
| Destination wrapper | +0xF8 | Wrapper embeds copied record at +8 |
| Subsequent transaction | +0x198 | Input RDX at OART +0x89CA6B |

Before transfer, the wrapper discriminator at +0 was 2. OART +0x22BCF4 calls
+0xA0870 with destination+8 and the loaded source, then changes discriminator
bits with `(value & ~6) | 1`. This is an observed property-record transfer, not a
standalone image-to-Shape API. Adjacent fields include other state and references
whose ownership/semantics have not been recovered.

The transaction constructor at OART +0x48870 writes through +0x508. The probe
therefore inspects only its known 0x510-byte stack region; it does not scan arbitrary
heap objects. An earlier 0x100-byte observation did not reach the cached pointer.
Absence in that smaller observation was not evidence that the transaction lacked it.

## Resolved dispatch path

```text
UserPicture loader returns                   OART +0x89C9F6
loaded-record transfer                      OART +0x22BCF4
transaction construction                    OART +0x48870
receiver vtable +0x78 call                   OART +0x89CA6B
    receiver vtable                         OART +0x9F6658
    resolved forwarding method              OART +0x21DEC0
    forwards unchanged to receiver slot +0x50
        resolved method                     OART +0x21BB20
        slot +0x60 call                     OART +0x21BBE8
            resolved preparation candidate  OART +0x2290F0
        slot +0x58 call                     OART +0x21BC06
            resolved application candidate  OART +0x1B88B0
returns to UserPicture                      OART +0x89CA71
```

This sequence was dynamically repeated. At the slot +0x58 call, RDX points to a
different operation object whose first qword is vtable OART +0x9E4BD0. It is not
the cached GFX image pointer. The later RAX value equals this operation address
in the observed run, but the public meaning and lifetime of that return are not
established; do not treat it as an HRESULT or retained result.

Static inspection of +0x2290F0 shows type-dependent branches, output storage in
R8, and possible construction using transaction+0x18. The exact branch taken
and operation constructor must be proved dynamically before assigning an ABI.
+0x1B88B0 changes operation flags, consults receiver methods and calls +0x1B8F50.
These descriptive aliases are hypotheses about roles, not recovered symbol names.

## Validation and remaining work

The harness now compares Shape ID, name, type, bounds, rotation and Z order before
and after tracing. The repeated run preserved those values and produced a picture
fill. The saved package contains one normal Shape, no Picture shapes and one PNG
matching the trace source's recorded SHA256. This validates ordinary UserPicture
under observation, not a synthesized transaction or optimized backend.

Next experiment: observe which preparation branch constructs the operation with
vtable +0x9E4BD0, determine which Shape/document identity it captures, and follow
its execution/cleanup through +0x1B8F50. Establish lifetime of its non-image state
before attempting reuse. The known GFX pointer at transaction+0x198 alone is not
a safe fill-binding interface.

No performance claim, production capability, or compatibility profile enabling
private calls changed. File-free loading, donor-free assignment, independent
resource lifetime and repeated decode avoidance are still unimplemented.

Raw evidence: `docs/evidence/fill_transaction_validated.txt` and
`docs/evidence/decoder_shape_validation.txt`. Saved package details:
`docs/fill_transaction_package.md`. Ignore nearest-export names in GDB stacks;
use verified module RVAs.
