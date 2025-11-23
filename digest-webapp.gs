/**
 * ==================== 다이제스트 웹앱 ====================
 * HTML 다이제스트를 웹으로 서빙 (카톡 미리보기용)
 */

/**
 * 웹앱 진입점
 * URL: https://script.google.com/.../exec?date=2025-11-21
 */
function doGet(e) {
  const params = e.parameter;
  const dateStr = params.date || getYesterdayDate();

  try {
    const htmlContent = 다이제스트HTML가져오기(dateStr);

    if (!htmlContent) {
      return HtmlService.createHtmlOutput(`
        <!DOCTYPE html>
        <html lang="ko">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>다이제스트 없음</title>
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
              display: flex;
              justify-content: center;
              align-items: center;
              height: 100vh;
              margin: 0;
              background: #f8f9fa;
            }
            .message {
              text-align: center;
              padding: 40px;
              background: white;
              border-radius: 12px;
              box-shadow: 0 2px 12px rgba(0,0,0,0.1);
            }
            h1 { color: #e74c3c; margin-bottom: 10px; }
            p { color: #7f8c8d; }
          </style>
        </head>
        <body>
          <div class="message">
            <h1>❌ 다이제스트를 찾을 수 없습니다</h1>
            <p>${dateStr} 날짜의 다이제스트가 없습니다.</p>
            <p style="font-size: 14px; margin-top: 20px;">
              URL 형식: <code>...exec?date=2025-11-21</code>
            </p>
          </div>
        </body>
        </html>
      `);
    }

    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle(`📚 ${dateStr} 스터디 다이제스트`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log(`웹앱 오류: ${error.message}`);

    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html lang="ko">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>오류</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
            background: #f8f9fa;
          }
          .error {
            text-align: center;
            padding: 40px;
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.1);
          }
          h1 { color: #e74c3c; margin-bottom: 10px; }
          p { color: #7f8c8d; }
          code {
            display: block;
            background: #f4f4f4;
            padding: 10px;
            border-radius: 4px;
            margin-top: 10px;
            font-size: 12px;
            color: #e74c3c;
          }
        </style>
      </head>
      <body>
        <div class="error">
          <h1>⚠️ 오류 발생</h1>
          <p>다이제스트를 불러오는 중 오류가 발생했습니다.</p>
          <code>${error.message}</code>
        </div>
      </body>
      </html>
    `);
  }
}

/**
 * 어제 날짜 가져오기
 */
function getYesterdayDate() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  return Utilities.formatDate(yesterday, 'Asia/Seoul', 'yyyy-MM-dd');
}

/**
 * 저장된 HTML 다이제스트 파일 가져오기
 * @param {string} dateStr - 날짜 (yyyy-MM-dd)
 * @returns {string} HTML 내용
 */
function 다이제스트HTML가져오기(dateStr) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
    const htmlFileName = `digest-${dateStr}.html`;

    const files = folder.getFilesByName(htmlFileName);

    if (!files.hasNext()) {
      Logger.log(`HTML 파일 없음: ${htmlFileName}`);
      return null;
    }

    const file = files.next();
    const htmlContent = file.getBlob().getDataAsString('UTF-8');

    return htmlContent;

  } catch (error) {
    Logger.log(`HTML 파일 읽기 실패: ${error.message}`);
    throw error;
  }
}

/**
 * 웹앱 URL 생성 도우미
 * 다이제스트 생성 후 이 함수를 호출하면 공유 가능한 URL을 얻을 수 있음
 */
function 웹앱URL생성(dateStr) {
  // 이 URL은 Apps Script 배포 후 자동으로 생성됨
  // 배포 > 웹 앱으로 배포 > URL 복사
  const WEB_APP_URL = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';

  if (!dateStr) {
    dateStr = getYesterdayDate();
  }

  const url = `${WEB_APP_URL}?date=${dateStr}`;

  Logger.log(`\n📱 카톡 공유 URL:`);
  Logger.log(url);

  return url;
}
