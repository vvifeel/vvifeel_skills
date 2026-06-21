# vvifeel_skills — 프로젝트 지침

## 버전 표기 규칙

스킬 내용(SKILL.md, references/, scripts/)을 수정할 때마다 반드시 아래 두 곳을 함께 업데이트한다.

### 1. SKILL.md description 필드 — `[vX.Y]` 태그

```yaml
description: >
  [v1.3] 스킬 설명 ...
```

- 기능 추가·동작 변경·참조 파일 개정 → Y +1 (minor bump)
- 오탈자·링크 등 내용 무관 수정 → Y 유지 가능
- 스킬 신규 생성 시 `[v1.0]`부터 시작

### 2. plugin.json — `"version"` semver

```json
"version": "0.4.0"
```

- 플러그인 내 스킬 중 하나라도 변경되면 minor 버전 +1 (0.3.0 → 0.4.0)
- 버그픽스·오탈자 등 동작 무관 수정만이면 patch +1 (0.4.0 → 0.4.1)
- 경로: `plugins/vertical-plugins/<plugin>/.claude-plugin/plugin.json`

### 체크리스트 (스킬 수정 PR마다 확인)

```
□ 변경된 스킬 SKILL.md description [vX.Y] 태그 bump
□ 해당 플러그인 plugin.json version bump
□ git commit 메시지에 버전 명시 (예: "feat: ... (ss-report-winword v1.3, plugin 0.4.0)")
```
