// Apps Script doGet() 함수 - 올바른 버전
// 기존 월별출석집계() 함수를 호출하는 방식 유지

function doGet(e) {
  const month = e.parameter.month;
  const type = e.parameter.type;

  Logger.log('doGet 호출: month=' + month + ', type=' + type);

  try {
    if (type === 'weekly') {
      // 🆕 주간 집계 데이터 반환 (JSON 파일에서 읽기)
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
      // 기존 일간 출석 데이터 반환 (JSON 파일에서 읽기)
      const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
      const fileName = `attendance_summary_${month}.json`;
      const files = folder.getFilesByName(fileName);

      if (!files.hasNext()) {
        return ContentService
          .createTextOutput(JSON.stringify({
            error: 'Attendance summary not found',
            message: '일간 출석 파일이 없습니다.'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const file = files.next();
      const content = file.getBlob().getDataAsString();

      return ContentService
        .createTextOutput(content)
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
