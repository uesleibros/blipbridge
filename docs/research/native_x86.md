# Native x86: research notes

**Status: reconnaissance. Nothing here is implemented, and nothing here has been
called.** These are observations from read-only probing of a real 32-bit
PowerPoint, recorded so the next session starts from evidence rather than from
memory. Where something is a hypothesis it is labelled as one.

## The host these observations come from

| | |
|---|---|
| Product | Microsoft PowerPoint, ProPlus2021Volume |
| Platform | **x86** (Click-to-Run `Platform=x86`) |
| POWERPNT.EXE | `C:\Program Files (x86)\Microsoft Office\Root\Office16\POWERPNT.EXE`, PE machine x86, WOW64 process confirmed via `IsWow64Process` |
| Binary build | **16.0.14334.20848** - see the correction below |
| Windows | 10.0.26200 |

Module identities as the compatibility framework reads them, from the live
process:

```
POWERPNT.EXE  x86 16.0.14334.20848 ts=0x6a73a611 size=0x1c9000  {817AC0EA-696D-439E-93AD-A982CDC88D66}+2
ppcore.dll    x86 16.0.14334.20848 ts=0x6a73a5e1 size=0x1508000 {1EE1FE1D-E316-46CC-BAFA-4B9BD20A1EBD}+2
oart.dll      x86 16.0.14334.20848 ts=0x6a73a340 size=0xbc8000  {94E0FD31-972C-477A-AFDD-F90A31B01378}+2
gfx.dll       x86 16.0.14334.20848 ts=0x6a73a532 size=0x40a000  {0F773FF9-7D9E-40E2-970A-A77F46A724B9}+2
mso.dll       x86 16.0.14334.20848 ts=0x6a73a5cf size=0x1851000 {5EA7220F-817B-4D19-802D-372268323804}+2
```

All six carry a build signature, so all are strong enough to authorise a profile.

### Correction: the build number, and why the first reading was wrong

An earlier note here recorded the build as **.20906**. That was wrong, and how it
was wrong is worth keeping.

`.20906` is the **Click-to-Run package version**
(`HKLM\...\ClickToRun\Configuration\ClientVersionToReport`). It is not the
version of any binary. The modules' own `VS_FIXEDFILEINFO` says `.20848`, and
that is what a profile is keyed to.

The two also disagreed *over time*. At 12:19 the files on disk reported `.20906`;
at 12:37 Click-to-Run rewrote them, and they now report `.20848`. Same paths,
same session, different binaries - which is a live demonstration of the thing
module identity exists for, and of why a resolved profile is re-checked against
the modules on every process start rather than trusted from a previous run.

Consequence for Phase C: the x86 binaries are **build .20848**, the same build
number the accelerated x64 backend was derived against. That is better ground for
an eventual same-build comparison than the earlier note suggested - but the build
*number* matching is not the same as the binaries matching, and the x64 profile's
own identity fields will have to be compared when an x64 install exists again.

## Method

`tools/probe_fill_structure.ps1` drives `InspectFillStructure`
(`experiments/exp_native_x86/structure_probe.cpp`), which performs **guarded
reads only**: `VirtualQuery` for commitment before every access, no calls into
anything it finds, no document touched beyond a disposable presentation with one
disposable AutoShape.

The probe is architecture-neutral by construction - the same source runs on x64
and x86 - so a difference in its output is a difference in Office rather than in
the instrument.

## Observation 1: `Shape.Fill` is still a PPCORE object

```
pointerSize = 4
object      = 0x10642678 (heap)
vtable      = ppcore.dll+0x1124404
```

Structurally the same starting point as x64: the object PowerPoint hands back
from `Shape.Fill` is a PPCORE wrapper, not an OART object directly.

## Observation 2: an OART object hangs off the wrapper at +0x4

```
+0x00  ppcore.dll+0x1124404        the wrapper's own vtable
+0x04  heap, [0] = oart.dll+0x90e7e8
+0x08  ppcore.dll+0xfaa5a4
+0x14  heap, [0] = ppcore.dll+0xfc8838
+0x20  ppcore.dll+0x13a74a4        a second vtable - a subobject
+0x24  heap, [0] = ppcore.dll+0xfaa5a4
+0x2c  points back at the object itself
```

The word at **+0x4** points to a heap object whose own vtable is in `oart.dll`.
On x64 the inner OART FillFormat sat at **+0x8**, one pointer in. Here it is one
pointer in as well.

**Hypothesis, not yet established:** the x86 wrapper holds the OART FillFormat at
`+0x4` for the same reason the x64 one holds it at `+0x8`, and the inner offset
is simply pointer-scaled. This is suggestive and is *not* proof - the vtable at
`oart.dll+0x90e7e8` has not been identified as a FillFormat, only as something
OART owns. Confirming it means checking what that object's vtable slots do, which
is the next step.

The self-reference at `+0x2c` and the second vtable at `+0x20` look like a
multiple-inheritance COM layout, which would be unremarkable.

## Observation 3: the wrapper's methods are **not** identity thunks

This is the finding that most clearly breaks the analogy with x64.

On x64 the wrapper's vtable was almost entirely identity thunks: slot *N* loaded
`this + innerOffset` and jumped to the inner object's slot *N*, encoded as
`48 8b 49 XX` and recognisable as a shape rather than as an address. That shape
is what made the x64 walk safe to derive.

On x86, every slot examined begins `55 8b ec` - `push ebp; mov ebp, esp` - a real
stack frame. These are functions, not tail-call thunks.

```
[ 0] ppcore.dll+0x4e81fa   55 8b ec 8b 4d 08 ff 75 10 8b 55 0c
[ 1] ppcore.dll+0x4e81de   55 8b ec 56 8b 75 08 8b 46 24 50 8b
[ 3] ppcore.dll+0x71e69c   55 8b ec 8b 4d 08 8b 55 0c 8d 49 20
```

Reading them:

* **Slot 0** - `mov ecx, [ebp+8]` takes the **first stack argument** into ECX,
  then pushes `[ebp+0x10]` and loads `[ebp+0xc]`. `this` arriving on the stack
  rather than in ECX is the COM convention on x86: interface methods are
  `__stdcall`, with `this` as the first stack parameter.
* **Slot 1** - `mov esi, [ebp+8]` (`this`), then `mov eax, [esi+0x24]` and pushes
  it. It forwards through a *field*, not through a fixed `this + k` adjustment.
* **Slot 3** - `mov ecx, [ebp+8]` (`this`), then `lea ecx, [ecx+0x20]` and calls.
  ECX-based, so the **callee** is `__thiscall` even though the wrapper method
  itself is `__stdcall`. `+0x20` matches the second vtable seen in the object
  words.

**Established:** the outer wrapper methods are `__stdcall` with `this` on the
stack. **Established:** at least one of them calls an inner method `__thiscall`
with `this` in ECX. **Not established:** which slot corresponds to the fill
operation, or whether any equivalent of the x64 receiver chain exists at all.

### Why this matters for the design

The x64 backend recognises the wrapper *structurally* - "a vtable of N identity
thunks with a consistent inner offset" - rather than by a hard-coded vtable
address. That recogniser cannot be reused: there are no identity thunks here to
count. An x86 recogniser will need a different structural invariant, and finding
one is a prerequisite for the adaptive resolver, not a detail of it.

## Observation 4: the object at `+0x4` is a COM object, and its vtable says so

Its first three slots are `IUnknown`, read from the prologues:

```
[0] oart.dll+0x1c3ad3  55 8b ec 57 8b 7d 10 85 ff 0f 84 cf
                       push ebp; mov ebp,esp; push edi
                       mov edi,[ebp+0x10]     ; third stack argument
                       test edi,edi; jz ...   ; ppvObject != null
[1] oart.dll+0x19c8a4  55 8b ec 8b 4d 08 8b 41 18 40 89 41
                       mov ecx,[ebp+8]        ; this, from the stack
                       mov eax,[ecx+0x18]; inc eax; mov [ecx+0x18],eax
[2] oart.dll+0x1ac901  55 8b ec 8b 4d 08 56 83 69 18 01 8b
                       mov ecx,[ebp+8]        ; this, from the stack
                       sub dword [ecx+0x18],1
```

**Established:**

* Slots 0/1/2 are `QueryInterface` / `AddRef` / `Release`. Slot 0 null-checks its
  third stack argument, which is `ppvObject`; slots 1 and 2 increment and
  decrement the same field.
* The convention is **`__stdcall`**: `this` arrives as the first stack argument
  (`[ebp+8]`), not in ECX.
* **The reference count is at `this+0x18`.**
* The object carries **two vtable pointers**, at `+0x0` (`oart.dll+0x90d888`) and
  `+0x4` (`oart.dll+0x90d84c`) - a multiple-inheritance layout, so it implements
  at least two interfaces.
* It holds a **back-pointer to the PPCORE Fill wrapper at `+0x8`**.

**Still not established:** which interface it is. "A reference-counted OART COM
object reachable from `Shape.Fill`, holding a back-pointer to its wrapper" is
what the evidence supports; it is not yet a FillFormat or anything else by name.

## Observation 5: the next object down uses a *different* calling convention

Reached at `+0x1c` from the object above; vtable `ppcore.dll+0xf50b4c`.

```
[0] ppcore.dll+0x2677ab  8b 41 0c 40 89 41 0c c3
                         mov eax,[ecx+0xc]; inc eax; mov [ecx+0xc],eax; ret
[1] ppcore.dll+0x48951f  56 8b f1 57 83 6e 0c 01
                         push esi; mov esi,ecx; push edi
                         sub dword [esi+0xc],1
[5] ppcore.dll+0x2af7e   8b 41 04 c3
                         mov eax,[ecx+4]; ret
```

**Established:**

* This object's methods are **`__thiscall`**: `this` is in **ECX**, and slot 0
  returns with a bare `ret` - no stack argument to clean.
* Its **reference count is at `this+0xc`**, not `+0x18`.
* It is therefore *not* the same kind of object as the one above it, and shares
  neither its convention nor its layout.

**This is the single most important result so far for correctness.** Two objects,
one hop apart in the same chain, use different calling conventions and keep their
reference counts at different offsets. Any model that assumed one convention for
"the x86 private objects" would be wrong for half of them, and the failure mode -
a stack imbalance on every call - is the kind that survives one successful call
and corrupts the process later.

It also confirms the rule the brief set: **the callee's convention must be read
per function, never inferred from the caller's.**

## Observation 6: where the picture fill actually lands

Found by walking every object reachable from `Shape.Fill` and diffing the walk
across documented `Fill.UserPicture` calls - `tools/probe_fill_graph_diff.ps1`.
Twelve objects are reachable at depth 3. The transitions:

| transition | objects mutated |
|---|---|
| no fill → picture A | the OART COM object (d1); an object with vtable `oart.dll+0x824198` (d3) |
| picture A → picture B *(different image and size)* | the `ppcore.dll+0xf50b4c` object (d2); the `oart.dll+0x824198` object (d3) |
| picture B → solid colour | the `oart.dll+0x824198` object (d3) |

**Established:**

* The object with vtable **`oart.dll+0x824198`** changes on *every* fill
  transition, including between two different pictures. It holds picture-specific
  state.
* The OART COM object changes when a fill *appears* but **not** between two
  different pictures - consistent with a type or dirty field rather than the
  image.
* Nothing at all changes in the PPCORE wrapper's or the OART object's first
  `0x80` bytes between two different pictures. An earlier reading that appeared
  to show the wrapper changing was **lazy initialisation on first touch**, not
  the fill; the harness now touches `Fill.Type` before the baseline so the two
  cannot be confused.

**Hypothesis:** the `oart.dll+0x824198` object is the picture-fill record, or
holds it. Not established - only that it is picture-dependent.

### Why the path to it is not a fixed chain

The intermediate object (`ppcore.dll+0xf50b4c`) looks like a **property list
rather than a struct**. Its fields repeat in a pattern:

```
+0x18 ptr  +0x1c ptr  +0x20 0x3dc  +0x24 ptr
+0x48 ptr  +0x4c ptr  +0x50 0x3dd  +0x54 ptr
+0x60 ...  +0x64 1    +0x68 1
```

Two parallel records with what look like consecutive property ids (`0x3dc`,
`0x3dd`). Dumping a fixed chain `4,0x1c,0x54` on a *second* presentation reached a
different object entirely, whose "vtable" was outside every known module - so
**`+0x54` is an entry in a list whose position depends on which properties exist**,
not a stable field offset.

Consequence for the resolver: the x86 invariant cannot be a fixed offset chain
from the wrapper. It will have to be structural - recognise the property list,
then find the entry by its id - which is a different shape of invariant from
x64's, exactly as the framework was built to allow.

## Correction: code addresses are not aligned on x86

The framework's structural validation required candidate code addresses to be
pointer-aligned. That is true enough on x64 and **false on x86**: real PPCORE
vtable slots point at `ppcore.dll+0x4e81fa`, `+0x73fd1e` and similar, none of
them 4-byte aligned. The compiler aligns hot functions and leaves the rest where
they fall.

The check was rejecting valid vtables, which is the expensive direction of wrong
- it would have sent the resolver to the portable backend on a build it could
have accelerated. Alignment is now required only of vtables, which are data.

Found by pointing the probe at real Office objects and reading "vtable valid: NO"
against a vtable that was plainly fine.

## Calling conventions: current state of knowledge

| object / function | convention | `this` | refcount | how established |
|---|---|---|---|---|
| PPCORE `Shape.Fill` wrapper vtable | `__stdcall` | `[ebp+8]` | not seen | prologues, every slot examined |
| Inner call from wrapper slot 3 | `__thiscall` | ECX | - | `lea ecx,[ecx+0x20]` before the call |
| OART COM object at wrapper `+0x4` | `__stdcall` | `[ebp+8]` | `this+0x18` | `IUnknown` prologues (obs. 4) |
| `ppcore.dll+0xf50b4c` object at `+0x1c` | **`__thiscall`** | **ECX** | **`this+0xc`** | `mov eax,[ecx+0xc]` / bare `ret` (obs. 5) |
| The picture-bearing `oart.dll+0x824198` object | **unknown** | - | - | located, not yet decoded |
| OART private entry points | **unknown** - none located yet | | | |
| GFX exports | **unknown** - not examined on x86 | | | |

Two objects one hop apart use different conventions and keep their reference
counts at different offsets. Nothing may be inferred from a neighbour.

One successful call would not establish any of these. They will need the
repeated-call stress the v0.9 brief asks for, and they will need it before
anything is called in anger.

## What has not been done

Everything else. Specifically, none of the following has been started:

* decoding the picture-bearing `oart.dll+0x824198` object;
* finding the property id that selects the picture entry in the `+0x1c` list;
* locating anything equivalent to the x64 receiver, its control block, or the
  apply slot;
* locating the record constructor, the image sub-record, the transfer, the
  stretch holder, the transaction constructor or the destructors;
* re-deriving the property record's size and slot offsets for 4-byte pointers;
* re-establishing the reference-counting rules;
* confirming the GFX cached-image creator's x86 export and signature;
* rebuilding the Shape compatibility matrix, which **must not** be inherited from
  x64;
* any runtime execution of any of it.

## Established since: module identity is strong enough to key a profile

The compatibility framework (Phase A) now reads a **CodeView build signature** -
the PDB GUID and age from the debug directory - for every Office module, and
requires one before a profile may authorise private calls. Version, timestamp and
image size are three integers that could coincide; a linker-generated GUID per
build cannot.

It is read from the *mapped* image rather than the file, which is safe precisely
here and would not be for code: the debug directory sits in a read-only section
with no relocations, so the mapped bytes are the file's bytes. On x86 that is
emphatically not true of `.text`, where relocation rewrites absolute addresses
throughout - a hash of mapped x86 code would differ between processes and be
useless as an identity.

## Constraint on validating any of this

The machine now has **only** 32-bit Office. The x64 accelerated backend cannot be
run, benchmarked or regression-tested here at all, which means the v0.9
performance-parity comparison - x86 native against x64 native, same machine, same
build - is not currently measurable. See the freeze/blocker report for the
options.
