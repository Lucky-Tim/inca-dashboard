# 해빙 운영 플랫폼 — 세션 공통 규칙

## 브랜드 자산 (유일 원천)
- 로고: `design/brand/having-logo-full.png` (풀 워드마크) · `having-symbol-512.png` / `having-symbol-192.png` (심볼)
- 색·폰트·간격 토큰: `design/tokens.json` — **모든 디자인 값은 이 파일이 유일 원천.** 파생 설명서: `design/design-system.md`
- 핵심 규칙: 브랜드 레드 **#A40000** 최우선 · 웹 화면=Pretendard, 문서/덱=맑은 고딕 · 라이트 테마(뷰어 등 대외)/다크 테마(대시보드 등 내부) 분리, 상호 혼용 금지

## 캠페인·수치 규율
- 동네선물 캠페인 문구·수치의 원천: `design/캠페인팩트시트_동네선물3개월무료.md` — **여기 없는 수치·조건을 지어내지 말 것**
- 실적 수치는 실측 데이터만 사용, 목표치는 반드시 "(안)" 표기

## 파일·배포 규칙
- 이 폴더는 git 저장소(GitHub Pages 배포: lucky-tim.github.io/inca-dashboard). **git add는 파일 지정으로만** — `git add -A` 금지 (PII 파일 커밋 사고 방지)
- git commit/push는 사용자가 직접 실행 (Claude는 파일 수정까지만)
- `Code.js`는 Google Apps Script 서버 코드 — 로컬 수정 후 Apps Script 편집기에 붙여넣고 "배포 관리 → 새 버전"으로 반영해야 적용됨 ("새 배포" 금지: URL 바뀜)
- 핵심 화면: `index.html`(대시보드·내부) · `viewer.html`(경영진 뷰어·대외) · `viewer-mgr.html`(관리자)
- 오케스트레이션 인수인계: `_오케스트레이션_상태_20260702.md` 참조

## UI 수정 원칙
- 새 UI 요소 추가 전 기존 함수·스타일 패턴을 grep으로 확인하고 동일 스타일로 맞출 것
