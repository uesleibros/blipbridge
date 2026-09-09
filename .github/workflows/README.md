# CI

`windows.yml` builds Release and Debug with MinGW-w64 UCRT, runs CTest, and
checks that every export the VBA wrapper binds to by name is present.

**What CI does not do.** The Office integration suites in `tools/` need a live
PowerPoint on a validated Office build, which hosted runners do not have. They
are not run here and are not reported as passing. They run on a machine with
Office 16.0.14334.20848 and their transcripts are committed under
`docs/evidence/`, so the evidence travels with the repository even though the
automation cannot reproduce it.

The split is deliberate:

| Suite | Where | Needs Office |
|---|---|---|
| `abi_contract` | CI and locally | no - asserts fail-closed behaviour |
| `com_contract` | CI and locally | no |
| `tools/test_*.ps1` | validated machine only | yes |
