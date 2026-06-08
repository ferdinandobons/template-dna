# Security Policy

## Supported versions

BrandDocs is in **alpha** (`0.x`). Security fixes are applied to the latest
`main` and the most recent tagged release only.

| Version | Supported |
|---------|-----------|
| `0.1.x` | ✅        |
| `< 0.1` | ❌        |

## Reporting a vulnerability

Please **do not** open a public issue for security problems.

Instead, use GitHub's private
[**Report a vulnerability**](https://github.com/ferdinandobons/brand-docs/security/advisories/new)
flow (Security → Advisories), or email the maintainer at
`1bonsegnaferdinando@gmail.com` with:

- a description of the issue and its impact,
- steps to reproduce (a minimal **synthetic** template is ideal - never send
  real company files),
- any suggested remediation.

You can expect an acknowledgement within **7 days** and a status update within
**30 days**.

## Scope & handling notes

BrandDocs parses untrusted Office (OOXML) files. Areas of particular interest:

- **XML parsing**: XXE, billion-laughs / entity-expansion, zip-bomb style
  payloads in `.docx` / `.pptx` / `.xlsx`.
- **Path handling**: zip-slip / path traversal when unpacking template parts
  or writing a `brand-kit/`.
- **Formula / macro content**: preserved cell formulas and any embedded
  active content.

The engine is self-contained (`python-docx` / `python-pptx` / `openpyxl` +
`lxml`) with no cloud calls, which limits exposure, but template input is still
untrusted and handled accordingly.
