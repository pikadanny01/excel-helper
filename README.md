# 엑셀 헬퍼 v1.2.0 — 직원용 엑셀 검색 웹사이트

AI를 사용하지 않는 정적 웹사이트입니다. 직원들이 일반적인 한국어로 질문하면 연관 검색어를 찾아 도움말을 보여줍니다. 엑셀을 직접 조작하는 프로그램이 아니며, **질문을 무조건 이해하거나 모든 함수를 검증한 제품은 아닙니다.**

## v1.2.0 업데이트

- 예시 표에 **열(A, B, C…) / 행(1, 2, 3…) 머리글**을 표시하여 엑셀과 같은 셀 좌표를 확인할 수 있습니다.
- 2003년생·미참여 등 일부 질문의 연도·참여 상태를 추출하여 예시 표 및 FILTER / COUNTIFS 수식에 반영합니다.
- 성별+참여 상태, 점수 기준, 이용 횟수 기준, 참여연도+횟수+참여 상태 등 **명시된 템플릿**에서 복합 조건 수식을 생성합니다.
- 조건 수식 만들기 도구: 기준 열·값·데이터 마지막 행을 직접 지정하여 FILTER 또는 COUNTIFS 수식을 구성합니다.
- CSV / TSV 파일 로컬 미리보기: 브라우저에서만 최대 2MB와 25행·15열 미리보기를 처리합니다. **XLSX 직접 분석은 제공하지 않습니다.**
- 나만의 도움말 추가, 복사, 브라우저 저장, 관리자용 JSON 다운로드. **해당 브라우저에만 저장되며 직원 간 자동 공유는 제공하지 않습니다.**
- 도움말 '도움됐어요/부족해요' 평가 기록을 브라우저에 보관하고 JSON으로 내보낼 수 있습니다. 중앙 서버 전송·집계는 제공하지 않습니다.
- 148개 도움말, 411개 함수명 색인, 그중 127개 함수에 자체 요약·예시를 제공합니다. 색인 전체의 설명이 검증됐다는 뜻은 아닙니다.

## 사이트 배포용 파일

`index.html`, `app.js`, `search-core.js`, `example-visuals.js`, `styles.css`, `data.js`, `.nojekyll`

일곱 파일을 같은 폴더(저장소 최상위)에 두세요. ZIP 자체를 GitHub 저장소에 올리지 말고 **압축 해제한 파일**을 올립니다.

### 기존 GitHub Pages에서 업데이트

1. GitHub 저장소 `excel-helper`를 엽니다.
2. Add file → Upload files를 눌러 압축 해제한 7개 파일을 올립니다. 기존과 이름이 같으면 파일을 갱신합니다.
3. Commit changes를 눌러 저장합니다. 기존 설정(`main` / `/(root)`)은 유지합니다.
4. GitHub Pages 배포가 완료되면 기존 주소를 새로고침합니다.

### GitHub 아이디가 보이지 않는 무료 주소 — Cloudflare Pages (추천)

기존 GitHub 저장소를 그대로 연결할 수 있습니다. Cloudflare 계정을 만든 뒤 [Cloudflare Dashboard](https://dash.cloudflare.com/) → Workers & Pages → Create application → Pages → **Import an existing Git repository / Connect to Git** → GitHub 계정 권한 승인 → `excel-helper` 저장소 선택 → `main` 배포 → Framework preset 없음 → Build command `exit 0` 또는 빈값 → Build output directory `.`(저장소 루트) → Deploy 순으로 진행합니다.

배포가 끝나면 `<프로젝트이름>.pages.dev` 주소가 생성됩니다. 예를 들어 `excel-helper.pages.dev`를 원한다면 프로젝트 이름 `excel-helper`를 시도할 수 있으나 **이미 사용 중이면 다른 이름을 정해야 합니다**. 그 주소는 지금 바로 개설되거나 예약된 주소가 아닙니다. GitHub에 코드를 커밋하면 Cloudflare Pages에 연결된 사이트도 자동 업데이트될 수 있습니다.

`excel-helper.com`처럼 `.com`만으로 구성된 독립 도메인을 쓰려면 소유·구매한 도메인을 등록하고 DNS를 연결해야 합니다. 기존 GitHub Pages도 사용자 소유 도메인을 지원하지만 기본 github.io 주소에서는 개인 ID를 없앨 수 없습니다.

Cloudflare 공식 안내: https://developers.cloudflare.com/pages/get-started/git-integration/

## 편집 및 개발

`data.json`을 수정한 뒤 `python sync_data.py` 실행하여 `data.js`를 갱신합니다. `data.js`를 GitHub에 업로드해야 전 직원에게 공통 적용됩니다. 브라우저의 '나만의 도움말' 입력은 각 PC 로컬 저장이므로 공유용 관리 시스템이 아닙니다.

검색 테스트: `node --test tests/search.test.js tests/search_v12.test.js`

브라우저 테스트: `python tests/browser_smoke.py` 및 `python tests/browser_v12.py` (Chromium/Playwright 필요)

## 개인정보 및 보안

질문과 CSV 미리보기는 이 페이지 자체에서 별도 서버 API로 보내지 않습니다. 다만 방문 로그, 설치된 브라우저 확장 프로그램, CDN 또는 접근 경로에서 제공자의 정책이 적용될 수 있습니다. 공개 사이트에 비공개 업무자료나 개인정보를 게시하지 마세요. 예시 명단은 모두 가상입니다.

## 제한 사항

기존 데이터는 주로 2026년 Excel Microsoft 365/2016/2019/2021/2024를 기준으로 작성된 보조 도움말로, 실제 Excel 지원 여부는 플랫폼·업데이트 채널에 따라 차이가 있습니다. **AI가 아닌 정해진 규칙 기반**이므로 지원 템플릿 밖의 조건을 스스로 추론하지 못합니다. XLSX 파일 직접 분석, 실제 Excel 자동 수정, 모든 함수 상세 해설, 직원 간 실시간 피드백 동기화는 아직 제공하지 않습니다.

공식 함수 도움말: https://support.microsoft.com/ko-kr/excel/excel-functions-alphabetical
