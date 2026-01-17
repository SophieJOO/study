// ===== Apps Script Web App 수정 =====
// 기존 doGet() 함수를 찾아서 아래 코드로 교체하세요

/**
 * 웹 앱 엔드포인트 - 일간 출석 & 주간 집계 API
 */
function doGet(e) {
  const month = e.parameter.month; // '2025-11' 형식
  const type = e.parameter.type;   // 'weekly' 또는 undefined(일간 출석)

  try {
    if (type === 'weekly') {
      // 🆕 주간 집계 데이터 반환
      const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
      const fileName = `weekly_summary_${month}.json`;
      const files = folder.getFilesByName(fileName);

      if (!files.hasNext()) {
        return ContentService
          .createTextOutput(JSON.stringify({
            error: 'Weekly summary not found',
            message: '주간 집계 파일이 없습니다. 이번달주간집계() 함수를 실행하세요.'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const file = files.next();
      const content = file.getBlob().getDataAsString();

      return ContentService
        .createTextOutput(content)
        .setMimeType(ContentService.MimeType.JSON);

    } else {
      // 기존 일간 출석 데이터 반환
      const jsonData = 월별출석집계(month);

      return ContentService
        .createTextOutput(JSON.stringify(jsonData))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    Logger.log('doGet 에러: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
