# ss-report

Claude Code용 한국어 내부 보고서 생성 플러그인입니다.

이 플러그인은 세 가지 스킬을 포함합니다.

| 스킬 | Flow | 권장 용도 |
|------|------|---------|
| `ss-report-full` | 3단계 (목차 → 페이지 계획 → 생성) | 정식 보고서 |
| `ss-report-light` | 2단계 (목차+페이지 계획 통합 → 생성) | 초안, 빠른 테스트 |
| `ss-report-winword` | full + Word COM 레이아웃 검증 루프 | Windows + MS Word 환경 (최고 품질) |

스킬은 정형 DOCX 내부 보고서를 생성하며 다음 기준을 함께 관리합니다.

- 섹션 계층
- 한 줄 길이
- 표 배치
- 풋노트 배치
- 단락 여백
- 한국어 보고서 문체
- 번들 스크립트 기반 검증

## 설치

마켓플레이스를 추가합니다.

```bash
/plugin marketplace add https://github.com/vvifeel/vvifeel_skills.git
```

플러그인을 설치합니다.

```bash
/plugin install ss-report
```

설치 후 Claude Code를 재시작합니다.

## 업데이트

```bash
/plugin update ss-report
```

## 트리거 예시

- 보고서 만들어줘
- 워드로 뽑아줘
- 내부 보고서 작성해줘
- 검토 보고서 형태로 정리해줘

## 포함 스킬

```text
skills/ss-report-full/SKILL.md
skills/ss-report-full/references/
skills/ss-report-full/scripts/

skills/ss-report-light/SKILL.md
skills/ss-report-light/references/
skills/ss-report-light/scripts/

skills/ss-report-winword/SKILL.md       ← full + Word COM 검증 (Windows 전용)
skills/ss-report-winword/references/
skills/ss-report-winword/scripts/
```

## ss-report-winword 사전 조건

- Windows OS
- Microsoft Word 설치
- `pip install pywin32`

비Windows 환경에서는 `ss-report-full`을 사용하세요.
