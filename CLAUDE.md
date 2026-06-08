# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

제일기획(Cheil)용 **YouTube 성과 분석 웹앱** ("YouTube Stats Analyzer"). YouTube Data API v3를 브라우저에서 직접 호출해 채널/영상 지표와 댓글을 수집하고, Excel·Word·PPT 보고서로 다운로드한다. UI 텍스트와 코드 주석은 한국어 기준이다.

## 빌드 / 실행

빌드 단계·패키지 매니저·테스트가 **없는 순수 정적 클라이언트 앱**이다 (`package.json` 없음). 의존성은 모두 `index.html`의 CDN `<script>`로 로드된다.

- 로컬 실행: `index.html`을 브라우저로 직접 열거나, 정적 서버로 서빙
  - 예: `python -m http.server 8000` 후 `http://localhost:8000` 접속
- 배포: GitHub Pages (정적 호스팅). 별도 빌드 산출물 없이 소스가 그대로 배포된다.

## 아키텍처

3개 파일로 구성된다: `index.html` (구조/CDN 의존성), `style.css` (Neo-brutalism 스타일), `app.js` (전체 로직).

### app.js 구조 (섹션 번호가 파일에 주석으로 표기되어 있음)

1. **전역 상태 & UI 매핑** — `API_KEY`는 `localStorage`(`YT_API_KEY`)에 저장. `UI` 객체가 모든 DOM 요소를 한곳에 모음. 서버 없음 → 키는 사용자 브라우저에만 보관.
2. **공통 유틸** — `log()`(콘솔 UI 출력), 동적 입력행 추가/삭제, 탭/라디오/형식버튼 이벤트 바인딩(`initUI()`).
3. **출력 형식 레이어** — `getSelectedFormat(tabId)` → `saveData(sheetsObj, baseName, tabId, commentTexts)`가 형식에 따라 `saveExcel`/`saveWord`/`savePPT`로 분기. **세 탭 모두 이 단일 진입점을 통해 저장한다.** 4번째 인자 `commentTexts`가 있으면(탭2/탭3) 인사이트를 계산해 보고서에 삽입한다(탭1은 생략).
   - **2-2 댓글 인사이트** — `analyzeComments(texts)`가 ① 감성 댓글 건수(긍/부/중립), ② 극성별 단어 리스트(`sentimentWords.positive/negative/neutral`, 각 단어+빈도), ③ 키워드 TOP15를 반환. 감성 건수는 **KNU 한국어 감성사전**(`data/senti_dict.json`, 런타임 `fetch`로 1회 로드 후 `SENTI_MAP` 캐시; `ensureSentiDict()`)의 극성 점수 합산 + 부정어 반전(`commentScore()`)으로 계산. 단어 리스트는 토큰 각각의 극성(`lookupScore()`)으로 분류. 사전 로드 실패 시 `POSITIVE_WORDS`/`NEGATIVE_WORDS` 자체 사전으로 폴백. (워드클라우드 기능은 제거됨 — 표/리스트로만 표현.)
4. **파싱 헬퍼** — URL에서 채널/영상 ID 추출, `parseDurationType`(61초 이하=쇼츠), `formatDateKST`(KST 변환).
5. **YouTube API 래퍼** — `apiGet()`이 모든 호출의 단일 통로(키 주입·에러 처리). `getChannelVideos()`는 **모드별로 다른 엔드포인트를 쓴다**: `recent`는 playlistItems(쿼터 저렴), `date`는 search(쿼터 비쌈).
6. **탭 핸들러** — `t1Run`(채널 종합 지표), `t2Run`(단일 영상 댓글), `t3Run`(채널 내 다중 영상 댓글 대량 수집). 댓글 수집은 `scrapeVideoComments()` 공용 함수로 대댓글 페이징까지 처리.

### 핵심 데이터 흐름 (작업 시 반드시 이해할 것)

모든 탭은 **`sheetsObj`라는 공통 중간 포맷**을 만든다: `{ "시트명": [ {열: 값, ...}, ... ] }` 형태의 객체. 시트명에는 이모지 프리픽스(📊📈📋💬)가 붙는다. `saveExcel/saveWord/savePPT`는 모두 이 동일 구조를 입력으로 받으므로, **새 출력 형식을 추가하려면 이 시그니처(`sheetsObj`, `filename`)만 맞추면 된다.** Word/PPT 생성기는 시트명의 이모지를 정규식으로 제거(`.replace(/[📊📈📋💬]/g, '')`)한 뒤 제목으로 쓴다.

## 주의사항

- **API 쿼터 비용**: 영상별로 `videos`/`commentThreads`를 반복 호출하므로 탭 3(대량 댓글)은 쿼터를 빠르게 소모한다. 기간(`date`) 모드의 search 호출도 비싸다. 루프에 새 API 호출을 추가할 때 쿼터 영향을 고려할 것.
- **시트명 31자 제한**: Excel 시트명은 31자로 잘린다(`saveExcel`). 채널/영상명도 파일명 안전화를 위해 `[\/*?:<>|]` 제거 후 잘라 쓴다.
- CRLF/LF: 저장소가 LF↔CRLF 자동 변환 경고를 낸다(Windows 환경). 정상이며 무시 가능.
