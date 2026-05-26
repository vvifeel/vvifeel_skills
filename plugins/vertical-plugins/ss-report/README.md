# ss-report

Korean internal report generation plugin for Claude Code.

This plugin packages one skill:

```text
skills/ss-report
```

The skill generates structured Korean internal reports as DOCX files with strict rules for:

- section hierarchy
- line length
- table placement
- footnote placement
- spacing
- Korean business writing style
- validation through bundled scripts

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

## Trigger Examples

- 보고서 만들어줘
- 워드로 뽑아줘
- 내부 보고서 작성해줘
- 검토 보고서 형태로 정리해줘

## Included Skill

```text
skills/ss-report/SKILL.md
skills/ss-report/references/
skills/ss-report/scripts/
```
