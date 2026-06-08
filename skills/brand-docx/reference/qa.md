# QA

M1 implements deterministic L0 checks:

- profile schema validity
- every indexed role has a resolver
- no literal markdown leaks into generated text
- no detected demo instruction text remains in generated output

On top of L0, a two-stage visual audit runs when renderers are present and the
QA mode asks for it:

- **L1**: deterministic pixel proxies on the rendered PNGs (`visual.blank_page`,
  `visual.edge_bleed`, `visual.no_pages`), each a WARNING that never fails the
  gate by itself.
- **L2**: a `visual_manifest.json` (PNG paths + a profile-derived checklist) the
  orchestrator reads to make the qualitative judgement and drive a repair loop.
  The engine never calls a model.

It is gated by `--qa` (`fast` = L0 only; `auto` = L0 + L1; `deep` = L0 + L1 +
manifest; `strict` = deep plus visual gate errors) and degrades gracefully when
`soffice`/`pdftoppm` are absent in `auto`/`deep` (L0 plus INFO/WARNING findings).
`strict` fails when full render proof is unavailable. See
[visual-audit.md](visual-audit.md).
