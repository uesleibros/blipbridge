# Private OART entry points used by the native apply

Module `oart.dll`, tested Office build **16.0.14334.20848 x64**. Nothing here is
documented by Microsoft; every row was derived by disassembly and, where marked,
confirmed at run time. Calling convention is the Microsoft x64 one throughout
(RCX, RDX, R8, R9, then stack at `[rsp+0x20]`).

Every function below is **byte-verified before it is called**: the first sixteen
bytes at its RVA are compared against the bytes recorded when its ABI was
derived, and a single mismatch aborts the whole operation before anything is
constructed. The signatures live in
`experiments/exp_internal_blip/native_apply.cpp`.

## Entry points

| RVA | Role | Parameters | Returns | Evidence |
|---|---|---|---|---|
| +0x14110 | property record constructor | RCX = 0x4E8-byte storage | none used | static + live |
| +0x14F7B0 | mark every property slot unset | RCX = record | none | static + live |
| +0x14580 | image sub-record constructor | RCX = 0x200-byte storage | none used | static + live |
| +0x8F94C | install a cached image into an image sub-record | RCX = sub-record, RDX = `TCntPtr<ICachedImage>*` | none used | static + live |
| +0x22BCF4 | copy an image sub-record into a record slot | RCX = record+0x88, RDX = sub-record | slot | static + live |
| +0x158C40 | build the `{payload, descriptor}` holder | RCX = zeroed 16-byte holder, RDX = 16 source bytes | none used | static + live |
| +0x15AC70 | set a counted-value slot | RCX = slot, RDX = holder | slot | static + live |
| +0x48870 | transaction constructor | RCX = 0x510 storage, RDX = record, R8D = flags, R9B = bool, `[rsp+0x20]` = identifier | transaction | static + live |
| receiver vtable +0x78 | apply the transaction | RCX = receiver, RDX = transaction | not established | static + live |
| +0x8C388 | transaction destructor | RCX = transaction | none | static + live |
| +0x3D870 | image sub-record destructor | RCX = sub-record | none | static + live |
| +0xAF80 | property record destructor | RCX = record | none | static + live |
| +0xB210 | counted holder destructor | RCX = holder | none | static + live |

The apply's return value is deliberately ignored. In one traced run RAX happened
to equal the transient operation's address, which is not a documented result and
must not be read as an HRESULT.

## Sizes

| Object | Size | How it was established |
|---|---:|---|
| property record | 0x4E8 | `+0x14110` writes through +0x4E4; the operation embeds the record at +0x68 and its own fields resume at +0x550 |
| image sub-record | 0x200 | heap allocation of 0x200 at `OART +0xD9A62` immediately before `+0x14580`, and the caller then writes +0x1F8 |
| transaction | 0x510 | `+0x48870` writes +0x500, +0x504 and +0x508 |
| counted holder | 0x10 | `+0x158C40` allocates 0x10 for the payload and writes the pair at +0 and +8 |

The experiment allocates each with margin, because over-allocating a local costs
nothing while a short buffer would be memory corruption.

## Ownership

The cached GFX image is intrusively counted: a 32-bit count at object+8,
AddRef through vtable slot 0, Release through vtable slot +8.

| Step | Effect on the cached image | Evidence |
|---|---|---|
| `GEL::ICachedImage::Create` | returns one owned reference (moved, not AddRef'd) | GFX +0x7680 prologue; measured `afterCreate = 1` |
| `+0x8F94C` | **AddRefs**; the caller keeps its own reference | `+0x8FB30` calls slot 0 before storing; measured 1 -> 2 |
| `+0x22BCF4` | AddRefs, through `+0xA0870` -> `+0x9848C` | measured 2 -> 3 |
| `+0x48870` | AddRefs, by copying the record into transaction+0x18 | measured 3 -> 4 |
| apply | takes further references into document state | measured 4 -> 7 |
| `+0x8C388` | releases the transaction's copy | measured 7 -> 6 |
| `+0x3D870` | releases the sub-record's reference | measured 6 -> 5 |
| `+0xAF80` | releases the record's reference | measured 5 -> 4 |
| our own release | releases the reference `Create` returned | leaves 3 held by the document |

Every argument is **borrowed** for the duration of the call. No function above
takes ownership of a caller's buffer, and the three constructed objects are
destroyed by their paired destructors in reverse construction order, which is
the order `OART +0x89CA71..+0x89CABF` uses.

The receiver is borrowed too, and must never be cached: a Shape deleted through
public COM still passes every check in the resolution chain. It is re-resolved
before each apply, which costs four loads.

## Open: the apply takes one more reference than UserPicture

At the same point in the sequence, an ordinary `UserPicture` moved the count
3 -> 5 across the receiver call, while the native apply moves 4 -> 7. The base
differs by one for a known reason - the experiment deliberately keeps its own
`Create` reference alive for the whole call so it can sample counts, whereas the
file loader releases its local as soon as `+0x8F94C` has taken one. That
accounts for the different starting point but **not** for the different delta:
+2 against +3.

The two runs are not otherwise comparable: the traced `UserPicture` replaced an
existing picture fill, while the native apply replaced a solid fill. Replacing
versus establishing a fill plausibly differs by one reference, but that has not
been measured.

Per the project's own rule, this counts as an unexplained delta and blocks
productionization until it is either explained or removed. It does not block the
research experiment, which releases everything it created and leaves the document
holding Office's own references.

## Guarding

Beyond the byte signatures, every operation checks that `oart.dll`, `ppcore.dll`
and `gfx.dll` are all the tested build, and verifies the vtable of every object
it walks to. A build whose layout differs fails at the first check rather than
executing against a wrong ABI.

`GEL::ICachedImage::Create` is not in the table above because it is an actual
GFX export resolved by name; its ABI is in `record_construction.md`.
