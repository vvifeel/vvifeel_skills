# vvifeel_skills

vvifeel 스킬을 배포하기 위한 Claude Code 플러그인 마켓플레이스입니다.

## 빠른 시작

Claude Code에서 마켓플레이스를 추가합니다.

```bash
/plugin marketplace add https://github.com/vvifeel/vvifeel_skills.git
```

보고서 플러그인을 설치합니다.

```bash
/plugin install ss-report
```

설치 후 Claude Code를 재시작합니다.

## 플러그인

| Plugin | 설명 |
| --- | --- |
| [ss-report](./plugins/vertical-plugins/ss-report) | 한국어 내부 보고서를 정형 DOCX 레이아웃, 표, 풋노트, 문체 검증 기준에 맞춰 생성하는 스킬 (ss-report-full / ss-report-light / ss-report-winword) |

## 요구사항

- Claude Code 플러그인 지원
- DOCX 생성 스크립트 실행을 위한 Node.js
- 후처리 및 검증 스크립트 실행을 위한 Python

## 라이선스

MIT
