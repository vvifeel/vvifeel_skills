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

| Plugin | Skills | Description |
| --- | --- | --- |
| [ss-report](./plugins/vertical-plugins/ss-report) | ss-report-full, ss-report-light | Korean internal report generation. Full-featured (3-step flow) and lightweight (2-step flow) variants for structured DOCX internal reports. |

## Requirements

- Claude Code plugin support
- Node.js for DOCX generation scripts used by the skill
- Python for post-processing and verification scripts

## License

MIT
