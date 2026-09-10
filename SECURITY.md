# Security policy

## What this project is, in security terms

BlipBridge calls **undocumented internal functions of Microsoft Office**. That is
its whole purpose, and it shapes what a security issue here looks like.

It does not open sockets, read files it was not given, execute downloaded code,
or elevate. The realistic failure modes are memory safety and data integrity
inside the host process: a Shape that reaches the private OART apply when it
should not, a stale pointer, a reference count that goes wrong, or a document
left corrupt.

One such bug has already been found and fixed - a Connector passed every
structural check and terminated PowerPoint - so these are not hypothetical. See
[docs/safety_model.md](docs/safety_model.md).

## Reporting a vulnerability

Use **GitHub's private vulnerability reporting** on this repository:
*Security → Report a vulnerability*. That keeps the report private until there is
a fix.

Please do not open a public issue for anything that could corrupt or crash a
user's documents until it has been looked at.

### What helps

* BlipBridge version (`BlipBridge.Version`) and ABI version.
* PowerPoint version and build, and whether Office is 32-bit or 64-bit.
* Windows version.
* `BlipBridge.Capabilities` output.
* The smallest sequence of calls that reproduces it.
* Whether the host survived, and whether a document was left damaged.

**Please do not attach memory dumps.** They can contain the contents of whatever
document was open. A description and a reproduction are more useful and carry far
less of your data. If a dump genuinely turns out to be necessary, that will be
discussed in the private report first.

### What is in scope

* Anything that lets a Shape class reach the private OART apply when the semantic
  policy should have refused it.
* Memory corruption, use-after-free, or a retained receiver pointer.
* Reference-counting errors that leak or over-release an Office object.
* Document corruption that survives save and reopen.
* A guard that can be bypassed - version checks, signature checks, vtable checks.

### What is not

* **Running on an unvalidated Office build.** The library refuses by design. If
  you find a way to make it *not* refuse, that is very much in scope.
* Crashes caused by patching the DLL or calling the C ABI from another thread -
  the threading contract is documented and enforced.
* Performance issues.

## Expectations

This is a single-maintainer project, not a funded product. Reports are taken
seriously and looked at as soon as is practical, but there is no response-time
commitment and no bounty.

## The guards, and why they are not optional

Anyone changing this code should know that the following exist for safety and are
not incidental:

* Office build and module version validation, with unknown builds refused.
* Structural validation of the PPCORE wrapper, the OART FillFormat, the control
  block and the receiver, each against a recorded layout.
* Byte-level signature checks on every private entry point before it is called.
* Semantic Shape eligibility, decided in one place, **before** any private call.
* Receiver re-resolution on every apply - a deleted Shape still passes every
  pointer check, so a cached receiver would be a use-after-free.
* Fail-closed behaviour on anything unrecognised.

`tools/test_semantic_guards.ps1` proves the dangerous refusals happen before the
private apply by counting entries into it. Weakening any of the above, or that
test, needs a much better reason than making something faster.
