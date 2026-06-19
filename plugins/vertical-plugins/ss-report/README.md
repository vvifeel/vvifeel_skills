# ss-report

Korean internal report generation plugin for Claude Code.

This plugin includes three skills:

| Skill | Flow | Best For |
|-------|------|----------|
| `ss-report-full` | 3-step (outline → page plan → generate) | Production reports |
| `ss-report-light` | 2-step (outline+page plan → generate) | Drafts, quick tests |
| `ss-report-winword` | full + Word COM layout verification loop | Windows + MS Word (highest quality) |

All skills generate structured Korean internal reports as DOCX files with strict rules for:

- Section hierarchy
- Line length
- Table placement
- Footnote placement
- Spacing
- Korean business writing style
- Validation through bundled scripts

## Install

Add the marketplace:

```bash
/plugin marketplace add https://github.com/vvifeel/vvifeel_skills.git
```

Install:

```bash
/plugin install ss-report
```

Restart Claude Code after installation.

## Update

```bash
/plugin update ss-report
```

## Trigger Examples

- 보고서 만들어줘
- 워드로 뽑아줘
- 내부 보고서 작성해줘
- 검토 보고서 형태로 정리해줘

## Included Skills

```text
skills/ss-report-full/SKILL.md
skills/ss-report-full/references/
skills/ss-report-full/scripts/

skills/ss-report-light/SKILL.md
skills/ss-report-light/references/
skills/ss-report-light/scripts/

skills/ss-report-winword/SKILL.md       ← full + Word COM verification (Windows only)
skills/ss-report-winword/references/
skills/ss-report-winword/scripts/
```

## Prerequisites for ss-report-winword

- Windows OS
- Microsoft Word installed
- `pip install pywin32`

On non-Windows environments, use `ss-report-full` instead.
