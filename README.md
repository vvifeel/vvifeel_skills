# vvifeel_skills

Claude Code plugin marketplace for vvifeel skills.

## Quick Start

Add this marketplace in Claude Code:

```bash
/plugin marketplace add https://github.com/vvifeel/vvifeel_skills.git
```

Install the report plugin:

```bash
/plugin install ss-report
```

Restart Claude Code after installation.

## Plugins

| Plugin | Description |
| --- | --- |
| [ss-report](./plugins/vertical-plugins/ss-report) | Korean internal report generation skill with strict DOCX layout, spacing, table, footnote, and writing-style validation. |
| [ss-report-light](./plugins/vertical-plugins/ss-report-light) | Korean internal report generation skill (lightweight) with streamlined 2-step flow. Generates structured DOCX internal reports with strict layout, spacing, table, footnote, and writing-style rules. |

## Requirements

- Claude Code plugin support
- Node.js for DOCX generation scripts used by the skill
- Python for post-processing and verification scripts

## License

MIT
