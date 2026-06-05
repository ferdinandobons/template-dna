# /brand-verify

Verify a saved Brand Profile. Reports QA findings and a verdict (the role map
lives in `PROFILE.md`; verify does not render a proof). Add `--accept` to mark a
passing profile as accepted.

Run from the plugin root (or set `TEMPLATE_DNA_ROOT` to it):

```bash
python scripts/brandkit/cli.py verify --name <brand> --scope auto --qa auto
```
