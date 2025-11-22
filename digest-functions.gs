/**
 * ==================== 일일 다이제스트 기능 ====================
 * 기존 apps script code.gs 파일에 추가할 코드
 *
 * 설치 방법:
 * 1. 이 파일의 내용을 복사
 * 2. "apps script code.gs" 파일 맨 아래에 붙여넣기
 * 3. 저장 후 트리거 설정
 */

/**
 * 매일 저녁 8시 실행 - 일일 다이제스트 생성
 * 트리거 설정: 매일 20:00
 */
function 일일다이제스트생성() {
  Logger.log('=== 일일 다이제스트 생성 시작 ===');

  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const digest = {};

  // 각 조원의 공부 내용 수집
  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];

    for (const folderId of folderIds) {
      const content = 공부내용추출(memberName, folderId, today);
      if (content) {
        digest[memberName] = content;
        break; // 첫 번째 폴더에서 찾으면 중단
      }
    }

    // 내용이 없으면 출석 상태 확인
    if (!digest[memberName]) {
      digest[memberName] = {
        상태: 출석상태확인(memberName, today),
        제목: null,
        요약: null,
        파일목록: [],
        썸네일: null
      };
    }
  }

  // JSON 파일로 저장
  다이제스트JSON저장(digest, today);

  // 카톡 공유용 메시지 생성
  const message = 카톡공유메시지생성(digest, today);

  // 방장에게 이메일 발송 (선택 사항 - 이메일 주소 변경 필요)
  // GmailApp.sendEmail('방장이메일@example.com', '[자동] 오늘의 스터디 다이제스트', message);

  Logger.log('=== 일일 다이제스트 생성 완료 ===');
  Logger.log(message);
}

/**
 * 공부 내용 추출 및 요약
 */
function 공부내용추출(memberName, folderId, dateStr) {
  try {
    const mainFolder = DriveApp.getFolderById(folderId);
    const subfolders = mainFolder.getFolders();

    // 오늘 날짜 폴더 찾기
    let targetFolder = null;
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      const folderName = folder.getName().trim();
      const dateInfo = 날짜추출(folderName);

      if (dateInfo && dateInfo.dateStr === dateStr) {
        targetFolder = folder;
        break;
      }
    }

    if (!targetFolder) return null;

    // 파일 분석
    const files = targetFolder.getFiles();
    const content = {
      상태: '출석',
      제목: '',
      요약: '',
      파일목록: [],
      썸네일: null,
      전체내용: '',
      폴더링크: targetFolder.getUrl()
    };

    let fileCount = 0;

    while (files.hasNext() && fileCount < 20) { // 최대 20개 파일만 처리
      const file = files.next();
      const fileName = file.getName();
      const mimeType = file.getMimeType();
      fileCount++;

      // OFF.md 체크
      if (fileName === 'OFF.md') {
        content.상태 = '오프';
        content.제목 = '오프';
        return content;
      }

      // 마크다운 파일 처리
      if (fileName.toLowerCase().endsWith('.md')) {
        try {
          const mdContent = file.getBlob().getDataAsString('UTF-8');
          content.전체내용 += mdContent + '\n\n';

          // 제목 추출 (# 으로 시작하는 첫 줄)
          if (!content.제목) {
            const titleMatch = mdContent.match(/^#\s+(.+)$/m);
            if (titleMatch) {
              content.제목 = titleMatch[1].trim();
            }
          }

          // 요약: 첫 300자 (공백 제거)
          if (!content.요약) {
            const cleanText = mdContent
              .replace(/^#.*$/gm, '') // 제목 제거
              .replace(/```[\s\S]*?```/g, '') // 코드블록 제거
              .replace(/!\[.*?\]\(.*?\)/g, '') // 이미지 제거
              .trim();

            if (cleanText.length > 0) {
              content.요약 = cleanText.substring(0, 300).trim();
              if (cleanText.length > 300) content.요약 += '...';
            }
          }

          content.파일목록.push({
            이름: fileName,
            타입: 'Markdown',
            링크: file.getUrl()
          });
        } catch (e) {
          Logger.log(`마크다운 파일 읽기 실패 (${fileName}): ${e.message}`);
        }
      }

      // PDF 파일 처리
      else if (mimeType === MimeType.PDF) {
        content.파일목록.push({
          이름: fileName,
          타입: 'PDF',
          링크: file.getUrl()
        });
      }

      // 이미지 파일 처리
      else if (mimeType.startsWith('image/')) {
        if (!content.썸네일) {
          // 이미지 썸네일 URL 생성
          content.썸네일 = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w400`;
        }

        content.파일목록.push({
          이름: fileName,
          타입: 'Image',
          링크: file.getUrl()
        });
      }

      // 기타 파일
      else {
        content.파일목록.push({
          이름: fileName,
          타입: 'File',
          링크: file.getUrl()
        });
      }
    }

    // 제목이 없으면 첫 파일명으로 대체
    if (!content.제목 && content.파일목록.length > 0) {
      content.제목 = content.파일목록[0].이름.replace(/\.(md|pdf|png|jpg|jpeg)$/i, '');
    }

    // 요약이 없으면 파일 목록으로 대체
    if (!content.요약 && content.파일목록.length > 0) {
      const fileNames = content.파일목록.map(f => f.이름).slice(0, 3);
      content.요약 = `${fileNames.join(', ')} 학습`;
      if (content.파일목록.length > 3) {
        content.요약 += ` 외 ${content.파일목록.length - 3}개`;
      }
    }

    return content;

  } catch (e) {
    Logger.log(`${memberName} 공부내용 추출 실패: ${e.message}`);
    return null;
  }
}

/**
 * 출석 상태 확인 (제출기록 시트에서 조회)
 */
function 출석상태확인(memberName, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) return '미확인';

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, recordDate, fileCount, links, folderLink, status, weekNum, reason] = data[i];

    const recordDateStr = typeof recordDate === 'string'
      ? recordDate
      : Utilities.formatDate(new Date(recordDate), 'Asia/Seoul', 'yyyy-MM-dd');

    if (name === memberName && recordDateStr === dateStr) {
      if (status === 'O') return '출석';
      if (status === 'OFF') return '오프';
      if (status === 'LONG_OFF') return '장기오프';
      if (status === 'X') return '결석';
    }
  }

  // 기록이 없으면 아직 제출 안함
  const now = new Date();
  const targetDate = new Date(dateStr);

  if (targetDate > now) {
    return '미래';
  } else {
    return '제출대기';
  }
}

/**
 * 다이제스트 JSON 파일로 저장 (웹페이지에서 사용)
 */
function 다이제스트JSON저장(digest, dateStr) {
  const jsonContent = JSON.stringify({
    date: dateStr,
    generated: new Date().toISOString(),
    members: digest
  }, null, 2);

  const fileName = `digest-${dateStr}.json`;
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  // 기존 파일 삭제
  const existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  // 새 파일 생성
  const file = folder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT);

  // 공개 설정 (웹에서 접근 가능하도록)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log(`다이제스트 JSON 저장 완료: ${fileName}`);
  Logger.log(`파일 URL: ${file.getUrl()}`);

  return file.getUrl();
}

/**
 * 카톡 공유용 메시지 생성
 */
function 카톡공유메시지생성(digest, dateStr) {
  let attendCount = 0;
  let offCount = 0;
  let absentCount = 0;
  const highlights = [];

  for (const [name, content] of Object.entries(digest)) {
    if (content.상태 === '출석') {
      attendCount++;

      // 하이라이트 수집 (파일 3개 이상, 긴 요약 등)
      if (content.파일목록.length >= 3) {
        highlights.push(`${name}님이 ${content.파일목록.length}개 파일 업로드`);
      }
      if (content.요약 && content.요약.length >= 200) {
        highlights.push(`${name}님의 ${content.제목 || '공부'} 정리가 상세함`);
      }
    } else if (content.상태 === '오프' || content.상태 === '장기오프') {
      offCount++;
    } else if (content.상태 === '결석') {
      absentCount++;
    }
  }

  // 메시지 구성
  let message = `📚 오늘의 스터디 다이제스트 (${dateStr})\n\n`;
  message += `✅ 출석: ${attendCount}명\n`;
  if (offCount > 0) message += `🏖️ 오프: ${offCount}명\n`;
  if (absentCount > 0) message += `❌ 결석: ${absentCount}명\n`;

  if (highlights.length > 0) {
    message += `\n🌟 하이라이트:\n`;
    highlights.slice(0, 3).forEach(h => {
      message += `• ${h}\n`;
    });
  }

  // 다이제스트 페이지 URL (실제 배포 시 변경 필요)
  const webAppUrl = ScriptApp.getService().getUrl();
  const digestUrl = `${webAppUrl}?page=digest&date=${dateStr}`;
  message += `\n🔗 자세히 보기: ${digestUrl}\n`;
  message += `\n💡 Tip: 링크를 클릭하면 조원들의 공부 내용을 예쁘게 볼 수 있어요!`;

  return message;
}

/**
 * 수동 실행용: 오늘의 다이제스트 메시지 확인
 */
function 오늘의다이제스트메시지확인() {
  일일다이제스트생성();
}

/**
 * 웹앱 진입점 - HTML 페이지 제공
 * doGet 함수가 이미 있다면 해당 함수를 수정하세요
 */
function doGet(e) {
  const page = e.parameter.page || 'attendance';
  const action = e.parameter.action;

  // 다이제스트 데이터 API
  if (action === 'getDigest') {
    const date = e.parameter.date || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    const digest = 저장된다이제스트불러오기(date);

    return ContentService
      .createTextOutput(JSON.stringify(digest))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 다이제스트 페이지
  if (page === 'digest') {
    return HtmlService.createHtmlOutputFromFile('digest-page')
      .setTitle('스터디 일일 다이제스트')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 기본: 기존 출석표
  return HtmlService.createHtmlOutputFromFile('index (1)')
    .setTitle('스터디 출석표')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 저장된 다이제스트 JSON 불러오기
 */
function 저장된다이제스트불러오기(dateStr) {
  const fileName = `digest-${dateStr}.json`;
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  try {
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      const content = file.getBlob().getDataAsString('UTF-8');
      return JSON.parse(content);
    }
  } catch (e) {
    Logger.log(`다이제스트 파일 읽기 실패: ${e.message}`);
  }

  // 파일이 없으면 즉시 생성
  Logger.log(`다이제스트 파일 없음. 즉시 생성: ${dateStr}`);
  일일다이제스트생성();

  // 재시도
  try {
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      const content = file.getBlob().getDataAsString('UTF-8');
      return JSON.parse(content);
    }
  } catch (e) {
    Logger.log(`다이제스트 재로드 실패: ${e.message}`);
  }

  // 실패 시 빈 데이터 반환
  return {
    date: dateStr,
    generated: new Date().toISOString(),
    members: {}
  };
}
