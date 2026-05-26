# ss-report

Claude Code용 한국어 내부 보고서 생성 플러그인입니다.

이 플러그인은 하나의 스킬을 포함합니다.

```text
skills/ss-report
```

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

## 트리거 예시

- 보고서 만들어줘
- 워드로 뽑아줘
- 내부 보고서 작성해줘
- 검토 보고서 형태로 정리해줘

## 포함 스킬

```text
skills/ss-report/SKILL.md
skills/ss-report/references/
skills/ss-report/scripts/
```
