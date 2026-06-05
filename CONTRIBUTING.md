# Contributing

Run the M1 smoke tests before submitting changes:

```bash
PYTHONPATH=scripts python -m unittest tests.test_smoke tests.test_no_proprietary
```

Never commit real customer `.docx`, `.pptx`, or `.xlsx` templates. Use synthetic
fixtures for tests.
