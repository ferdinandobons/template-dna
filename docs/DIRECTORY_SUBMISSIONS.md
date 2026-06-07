# Directory Submission Guide

Use this file as the reusable outreach kit for getting BrandDocs listed in
Claude Code, AI agent, document-generation, and open-source tool directories.

## One-line pitch

BrandDocs turns existing Word, PowerPoint and Excel templates into reusable AI
document-generation skills for Claude Code and Codex, generating faithful
on-brand `.docx`, `.pptx` and `.xlsx` files from the original template shell.

## Short listing copy

BrandDocs is an open-source Claude Code and Codex skill bundle for AI Office
document generation from real company templates. It extracts a reusable Brand
Profile from `.docx`, `.pptx` or `.xlsx` templates, then generates new documents
from the original shell while preserving brand styles, structure, layouts,
named ranges and Excel formulas. Unlike generic AI document generators,
BrandDocs is faithful by construction and runs locally.

## Suggested tags

`claude-code`, `codex`, `ai-agents`, `document-generation`,
`document-automation`, `office-automation`, `docx`, `pptx`, `xlsx`, `ooxml`,
`template-automation`, `brand-automation`, `python-docx`, `python-pptx`,
`openpyxl`

## Suggested listing fields

| Field | Value |
|---|---|
| Name | BrandDocs |
| Repository | `https://github.com/ferdinandobons/brand-docs` |
| Website | `https://ferdinandobons.github.io/brand-docs/` |
| Category | AI document generation / Claude Code plugin / Office automation |
| License | MIT |
| Formats | DOCX, PPTX, XLSX |
| Skills | `brand-docx`, `brand-pptx`, `brand-xlsx` |
| Install | `/plugin marketplace add ferdinandobons/brand-docs` then `/plugin install brand-docs` |

## Priority directories and lists

### Fast path

1. Pick one list from the targets below.
2. Read its contribution rules.
3. Fork it, add the short entry in the closest matching section, and open a PR.
4. Use the pull-request message template in this file.
5. If the list has no fitting section, open an issue first and ask maintainers
   whether "Claude Code plugins", "AI agent skills", "document generation", or
   "Office automation" is the best category.

### GitHub awesome lists

For GitHub-hosted lists, the usual process is:

1. Fork the directory/list repository.
2. Add BrandDocs to the most relevant section (`Plugins`, `Skills`,
   `Document generation`, `Office automation`, or `Claude Code`).
3. Keep the entry short and factual.
4. Open a pull request that explains why the project belongs there.

Suggested entry:

```markdown
- [BrandDocs](https://github.com/ferdinandobons/brand-docs) - Claude Code and
  Codex skill bundle that turns Word, PowerPoint and Excel templates into
  reusable AI document-generation skills for faithful `.docx`, `.pptx` and
  `.xlsx` output.
```

Good targets:

- [hesreallyhim/awesome-claude-code](https://github.com/hesreallyhim/awesome-claude-code)
- [ccplugins/awesome-claude-code-plugins](https://github.com/ccplugins/awesome-claude-code-plugins)
- [jeremylongshore/claude-code-plugins-plus-skills](https://github.com/jeremylongshore/claude-code-plugins-plus-skills)
- [github/explore](https://github.com/github/explore) topic pages if a relevant
  topic accepts examples

### Skill and plugin directories

Use each directory's submit button, issue template, or pull-request process.
Reuse the listing fields above.

Potential targets:

- Claude plugin directories and marketplaces
- Claude skill directories
- AI agent skill marketplaces
- MCP/tool marketplaces with document-generation categories
- Office automation directories

### Community launch posts

Good launch channels:

- Reddit: `r/ClaudeCode`, `r/ClaudeAI`, `r/LocalLLaMA`, `r/opensource`,
  `r/SideProject`
- Hacker News: `Show HN`
- LinkedIn: AI automation, operations, consulting, brand/comms, and knowledge
  work audiences
- GitHub Discussions in relevant Claude Code/plugin communities

## Pull request message template

```markdown
Hi! I would like to add BrandDocs to this list.

BrandDocs is an MIT-licensed Claude Code and Codex skill bundle for generating
on-brand Office documents from real company templates. It supports Word
(`.docx`), PowerPoint (`.pptx`) and Excel (`.xlsx`) through three skills:
`brand-docx`, `brand-pptx` and `brand-xlsx`.

What makes it distinct from generic document generators is that it extracts a
reusable Brand Profile from the original template, generates from the original
template shell, and runs deterministic QA to prevent style/layout/range drift.

Repo: https://github.com/ferdinandobons/brand-docs
Website: https://ferdinandobons.github.io/brand-docs/
License: MIT
```

## Ongoing checklist

- Add the website URL to the GitHub repository homepage field.
- Keep README and website title aligned with the primary keyword:
  "AI Office document generator from Word, PowerPoint and Excel templates".
- Publish a release once the alpha surface is stable enough to recommend.
- Repost major improvements with concrete examples rather than generic launch
  copy.
- Ask directory maintainers for the exact category if the list has no obvious
  "document generation" or "Office automation" section.
