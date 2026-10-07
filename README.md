# 수학 월별 레슨플랜

무료 공용 배포용 앱입니다. GitHub Pages는 접속 안내 페이지이며, 작성 화면과 데이터 저장은 Google Apps Script와 스프레드시트에서 처리합니다.

설정은 [SETUP.md](SETUP.md)를 참조하세요. Apps Script에는 `apps-script/Code.gs`와 `apps-script/Index.html` 두 파일을 사용합니다.
공용 접속 비밀번호와 스프레드시트 ID는 Google Script Properties에만 설정합니다. 코드에 비밀번호나 연결키를 넣지 마세요.

GitHub Pages는 `main` 브랜치 최상위 폴더에서 배포합니다. 실제 Google 앱이 검증된 뒤 `config.js`에 공개 웹 앱 URL을 입력합니다.

검증: `node tests/free-app.cjs`. Google 서비스를 모의 구현한 기능 검사이며 실제 Google 계정 배포와 브라우저 검증은 별도로 필요합니다.
