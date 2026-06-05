# /brand-generate

Generate an on-brand document from a saved Brand Profile.

DOCX/PPTX use an IntermediateDocument. XLSX uses a GridDocument fill manifest.

```bash
python scripts/cli.py generate --name <brand> --input idoc.json --output out.docx --scope auto --qa auto
```
