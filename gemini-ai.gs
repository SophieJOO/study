/**
 * ==================== 스터디 컨텐츠 수집 시스템 ====================
 * 공부 내용 자동 수집 및 정리
 */

/**
 * 마크다운 내용 클린업
 * HTML 태그, Obsidian 특수 문법 등을 제거하여 깔끔하게 정리
 * @param {string} content - 원본 마크다운 내용
 * @returns {string} 정리된 내용
 */
function 마크다운클린업(content) {
  if (!content) return '';

  let cleaned = content;

  // 1. HTML 태그 제거 (스타일 속성 포함)
  cleaned = cleaned.replace(/<span[^>]*>/g, '');
  cleaned = cleaned.replace(/<\/span>/g, '');
  cleaned = cleaned.replace(/<font[^>]*>/g, '');
  cleaned = cleaned.replace(/<\/font>/g, '');
  cleaned = cleaned.replace(/<[^>]+>/g, ''); // 나머지 HTML 태그 제거

  // 2. Obsidian 이미지 링크 변환
  cleaned = cleaned.replace(/!\[\[([^\]]+)\]\]/g, '[이미지: $1]');

  // 3. 연속된 빈 줄 정리 (3개 이상의 빈 줄을 2개로)
  cleaned = cleaned.replace(/\n{4,}/g, '\n\n\n');

  // 4. 마크다운 강조 문법은 유지 (**, *, ~~, ` 등)

  return cleaned.trim();
}

/**
 * 마크다운을 HTML로 변환
 * @param {string} markdown - 마크다운 텍스트
 * @returns {string} HTML
 */
function 마크다운을HTML로(markdown) {
  if (!markdown) return '';

  let html = markdown;

  // 1. 코드 블록 먼저 처리 (변환 전에 보호)
  const codeBlocks = [];
  html = html.replace(/`([^`]+)`/g, function(match, code) {
    codeBlocks.push(code);
    return `__CODE_BLOCK_${codeBlocks.length - 1}__`;
  });

  // 2. 제목 변환 (### → h3, ## → h2, # → h1)
  html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');

  // 3. 굵게 **text** → <strong>text</strong>
  html = html.replace(/\*\*([^\*]+)\*\*/g, '<strong>$1</strong>');

  // 4. 기울임 *text* → <em>text</em>
  html = html.replace(/\*([^\*]+)\*/g, '<em>$1</em>');

  // 5. 취소선 ~~text~~ → <del>text</del>
  html = html.replace(/~~([^~]+)~~/g, '<del>$1</del>');

  // 6. 리스트 변환
  const lines = html.split('\n');
  let inList = false;
  let result = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // 들여쓰기된 리스트 (탭 또는 2칸 공백)
    if (/^\t[\-\*] (.+)$/.test(line) || /^  [\-\*] (.+)$/.test(line)) {
      const content = line.replace(/^\t[\-\*] /, '').replace(/^  [\-\*] /, '');
      if (!inList) {
        result.push('<ul>');
        inList = true;
      }
      result.push(`  <li style="margin-left: 20px;">${content}</li>`);
    }
    // 일반 리스트
    else if (/^[\-\*] (.+)$/.test(line)) {
      const content = line.replace(/^[\-\*] /, '');
      if (!inList) {
        result.push('<ul>');
        inList = true;
      }
      result.push(`  <li>${content}</li>`);
    }
    // 리스트가 아닌 줄
    else {
      if (inList) {
        result.push('</ul>');
        inList = false;
      }
      result.push(line);
    }
  }

  // 마지막 리스트 닫기
  if (inList) {
    result.push('</ul>');
  }

  html = result.join('\n');

  // 7. 문단 처리
  html = html.replace(/\n\n+/g, '</p><p>');
  html = html.replace(/\n/g, '<br>\n');

  // h 태그와 ul 태그 주변의 불필요한 <br> 제거
  html = html.replace(/<br>\s*<\/h([123])>/g, '</h$1>');
  html = html.replace(/<h([123])><br>/g, '<h$1>');
  html = html.replace(/<br>\s*<ul>/g, '<ul>');
  html = html.replace(/<\/ul><br>/g, '</ul>');

  // 전체를 <p>로 감싸기
  if (!html.startsWith('<')) {
    html = '<p>' + html;
  }
  if (!html.endsWith('>')) {
    html = html + '</p>';
  }

  // 빈 <p></p> 제거
  html = html.replace(/<p><\/p>/g, '');
  html = html.replace(/<p>\s*<\/p>/g, '');
  html = html.replace(/<p><br><\/p>/g, '');

  // 8. 코드 블록 복원
  codeBlocks.forEach((code, index) => {
    html = html.replace(
      `__CODE_BLOCK_${index}__`,
      `<code>${code.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</code>`
    );
  });

  return html;
}


/**
 * 일일 다이제스트 생성
 * @param {string} dateStr - 날짜 (yyyy-MM-dd). 없으면 어제
 * @returns {string} 생성된 다이제스트
 */
function 일일AI다이제스트생성(dateStr) {
  if (!dateStr) {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    dateStr = Utilities.formatDate(yesterday, 'Asia/Seoul', 'yyyy-MM-dd');
  }

  Logger.log(`=== ${dateStr} 다이제스트 생성 시작 ===`);

  const 조원데이터 = [];

  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];

    Logger.log(`\n${memberName} 파일 찾는 중... (폴더 ID: ${folderIds.length}개)`);

    for (const folderId of folderIds) {
      const content = 파일내용수집(memberName, folderId, dateStr);

      if (content && content.내용) {
        Logger.log(`  ✅ 파일 발견! (${content.파일목록.length}개)`);

        조원데이터.push({
          이름: memberName,
          내용: content.내용,
          파일목록: content.파일목록
        });

        break; // 첫 번째 폴더에서 찾으면 중단
      } else {
        Logger.log(`  ❌ 이 폴더에서 찾을 수 없음`);
      }
    }
  }

  if (조원데이터.length === 0) {
    Logger.log('\n❌ 해당 날짜에 공부한 조원이 없습니다.');
    return null;
  }

  Logger.log(`\n✅ ${조원데이터.length}명의 데이터 수집 완료`);

  // 간단한 통합 요약 생성
  let 통합다이제스트 = `📚 ${dateStr} 스터디 다이제스트\n\n`;
  통합다이제스트 += `총 ${조원데이터.length}명 참여\n\n`;

  조원데이터.forEach((data, index) => {
    통합다이제스트 += `${index + 1}. ${data.이름} - ${data.파일목록.length}개 파일 제출\n`;
  });

  Logger.log('\n=== 다이제스트 완성 ===');
  Logger.log(통합다이제스트);

  다이제스트저장(통합다이제스트, 조원데이터, dateStr);

  return 통합다이제스트;
}

/**
 * 파일 내용 수집
 * @param {string} memberName - 조원 이름
 * @param {string} folderId - 조원 폴더 ID
 * @param {string} dateStr - 날짜 (yyyy-MM-dd)
 * @returns {Object} {내용, 파일목록}
 */
function 파일내용수집(memberName, folderId, dateStr) {
  try {
    // dateStr은 이미 yyyy-MM-dd 형식 (예: 2025-11-21)
    Logger.log(`  찾는 중: ${dateStr}`);

    // 조원 폴더
    const memberFolder = DriveApp.getFolderById(folderId);

    // 날짜 폴더 찾기 (yyyy-MM-dd 형식)
    const dateFolders = memberFolder.getFoldersByName(dateStr);
    if (!dateFolders.hasNext()) {
      Logger.log(`  날짜 폴더 없음: ${dateStr}`);
      return null;
    }

    const dateFolder = dateFolders.next();
    Logger.log(`  ✅ 폴더 발견: ${dateFolder.getName()}`);

    let 전체내용 = '';
    const 파일목록 = [];
    const files = dateFolder.getFiles();

    let fileCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const mimeType = file.getMimeType();

      fileCount++;

      // 마크다운 파일
      if (fileName.toLowerCase().endsWith('.md')) {
        try {
          const mdContent = file.getBlob().getDataAsString('UTF-8');

          // 마크다운 클린업 적용
          const cleanedContent = 마크다운클린업(mdContent);

          // 제목 추출
          const titleMatch = cleanedContent.match(/^#\s+(.+)$/m);
          if (titleMatch) {
            전체내용 += `[제목: ${titleMatch[1]}]\n\n`;
          }

          전체내용 += cleanedContent + '\n\n' + '='.repeat(50) + '\n\n';

          파일목록.push({
            이름: fileName,
            타입: 'Markdown'
          });

        } catch (e) {
          Logger.log(`  MD 파일 읽기 실패: ${fileName}`);
        }
      }

      // 텍스트 파일
      else if (mimeType === MimeType.PLAIN_TEXT || fileName.toLowerCase().endsWith('.txt')) {
        try {
          const txtContent = file.getBlob().getDataAsString('UTF-8');
          전체내용 += `[텍스트 파일: ${fileName}]\n\n${txtContent}\n\n` + '='.repeat(50) + '\n\n';

          파일목록.push({
            이름: fileName,
            타입: 'Text'
          });
        } catch (e) {
          Logger.log(`  텍스트 파일 읽기 실패: ${fileName}`);
        }
      }

      // PDF (파일명만 - OCR은 별도 구현 필요)
      else if (mimeType === MimeType.PDF) {
        전체내용 += `[PDF 문서: ${fileName}]\n\n`;
        파일목록.push({
          이름: fileName,
          타입: 'PDF'
        });
      }

      // 이미지
      else if (mimeType.startsWith('image/')) {
        전체내용 += `[이미지: ${fileName}]\n\n`;
        파일목록.push({
          이름: fileName,
          타입: 'Image'
        });
      }
    }

    Logger.log(`  총 ${fileCount}개 파일, 텍스트 추출: ${파일목록.length}개`);

    if (전체내용.trim().length > 0) {
      return { 내용: 전체내용, 파일목록 };
    }

    // 텍스트가 없어도 파일이 있으면 기본 정보 반환
    if (파일목록.length > 0) {
      return {
        내용: `${memberName}이(가) ${dateStr}에 ${파일목록.length}개 파일을 제출했습니다.`,
        파일목록
      };
    }

    return null;

  } catch (e) {
    Logger.log(`  ${memberName} 파일 수집 실패: ${e.message}`);
    Logger.log(`  Stack: ${e.stack}`);
    return null;
  }
}

/**
 * 다이제스트 저장
 * @param {string} 통합다이제스트 - 통합 다이제스트 텍스트
 * @param {Array} 조원데이터 - 조원별 상세 데이터
 * @param {string} dateStr - 날짜
 */
function 다이제스트저장(통합다이제스트, 조원데이터, dateStr) {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  // 1. 전체 원본 내용 파일 생성
  let 전체내용 = `📚 ${dateStr} 스터디 전체 내용\n`;
  전체내용 += `생성일시: ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')}\n`;
  전체내용 += `총 ${조원데이터.length}명 참여\n`;
  전체내용 += '='.repeat(80) + '\n\n';

  조원데이터.forEach((data, index) => {
    전체내용 += `\n${'#'.repeat(80)}\n`;
    전체내용 += `# ${index + 1}. ${data.이름}\n`;
    전체내용 += `${'#'.repeat(80)}\n\n`;

    전체내용 += `📁 제출 파일 (${data.파일목록.length}개):\n`;
    data.파일목록.forEach(file => {
      전체내용 += `  - ${file.이름} (${file.타입})\n`;
    });
    전체내용 += '\n';

    전체내용 += `📖 전체 내용:\n`;
    전체내용 += '-'.repeat(80) + '\n';
    전체내용 += data.내용 + '\n';
    전체내용 += '-'.repeat(80) + '\n\n';
  });

  const fullFileName = `full-content-${dateStr}.txt`;
  const existingFull = folder.getFilesByName(fullFileName);
  while (existingFull.hasNext()) {
    existingFull.next().setTrashed(true);
  }

  const fullFile = folder.createFile(fullFileName, 전체내용, MimeType.PLAIN_TEXT);
  fullFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 2. 간단 요약 파일 저장
  const summaryFileName = `summary-${dateStr}.txt`;
  const existingSummary = folder.getFilesByName(summaryFileName);
  while (existingSummary.hasNext()) {
    existingSummary.next().setTrashed(true);
  }

  const summaryFile = folder.createFile(summaryFileName, 통합다이제스트, MimeType.PLAIN_TEXT);
  summaryFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 3. JSON 데이터 저장
  const jsonData = {
    date: dateStr,
    generated: new Date().toISOString(),
    summary: 통합다이제스트,
    memberCount: 조원데이터.length,
    members: 조원데이터.map(data => ({
      name: data.이름,
      fileCount: data.파일목록.length,
      files: data.파일목록,
      fullContent: data.내용
    }))
  };

  const jsonFileName = `digest-${dateStr}.json`;
  const existingJson = folder.getFilesByName(jsonFileName);
  while (existingJson.hasNext()) {
    existingJson.next().setTrashed(true);
  }

  const jsonFile = folder.createFile(
    jsonFileName,
    JSON.stringify(jsonData, null, 2),
    MimeType.PLAIN_TEXT
  );
  jsonFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 4. HTML 파일 생성 (카톡 미리보기용)
  let htmlContent = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📚 ${dateStr} 스터디 다이제스트</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
            line-height: 1.7;
            color: #333;
            background: #f8f9fa;
            padding: 20px;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            margin-bottom: 40px;
            padding-bottom: 20px;
            border-bottom: 3px solid #4CAF50;
        }
        .header h1 {
            font-size: 28px;
            color: #2c3e50;
            margin-bottom: 10px;
        }
        .meta {
            color: #7f8c8d;
            font-size: 14px;
        }
        .member-section {
            margin-bottom: 50px;
            padding: 30px;
            background: #f8f9fa;
            border-radius: 8px;
            border-left: 4px solid #4CAF50;
        }
        .member-section h2 {
            font-size: 24px;
            color: #2c3e50;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
        }
        .member-section h2::before {
            content: "👤";
            margin-right: 10px;
        }
        .file-list {
            background: white;
            padding: 15px;
            border-radius: 6px;
            margin: 15px 0;
        }
        .file-list h3 {
            font-size: 16px;
            color: #34495e;
            margin-bottom: 10px;
        }
        .file-list ul {
            list-style: none;
            padding-left: 0;
        }
        .file-list li {
            padding: 8px 0;
            border-bottom: 1px solid #ecf0f1;
            color: #555;
        }
        .file-list li:last-child {
            border-bottom: none;
        }
        .file-list li::before {
            content: "📄";
            margin-right: 8px;
        }
        .content-section {
            margin-top: 20px;
        }
        .content-section h3 {
            font-size: 18px;
            color: #34495e;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ecf0f1;
        }
        .content-body {
            background: white;
            padding: 20px;
            border-radius: 6px;
            line-height: 1.8;
        }
        .content-body h1 {
            font-size: 22px;
            color: #2c3e50;
            margin: 25px 0 15px 0;
            padding-bottom: 10px;
            border-bottom: 2px solid #4CAF50;
        }
        .content-body h2 {
            font-size: 20px;
            color: #34495e;
            margin: 20px 0 12px 0;
        }
        .content-body h3 {
            font-size: 18px;
            color: #555;
            margin: 15px 0 10px 0;
        }
        .content-body ul {
            margin: 15px 0;
            padding-left: 25px;
        }
        .content-body li {
            margin: 8px 0;
        }
        .content-body strong {
            color: #2c3e50;
            font-weight: 600;
        }
        .content-body em {
            color: #7f8c8d;
            font-style: italic;
        }
        .content-body code {
            background: #f4f4f4;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: "Monaco", "Courier New", monospace;
            font-size: 0.9em;
            color: #e74c3c;
        }
        .content-body p {
            margin: 12px 0;
        }
        @media (max-width: 768px) {
            .container {
                padding: 20px;
            }
            .header h1 {
                font-size: 22px;
            }
            .member-section {
                padding: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📚 ${dateStr} 스터디 다이제스트</h1>
            <div class="meta">
                생성일시: ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')} |
                참여: ${조원데이터.length}명
            </div>
        </div>
`;

  조원데이터.forEach((data, index) => {
    htmlContent += `
        <div class="member-section">
            <h2>${data.이름}</h2>

            <div class="file-list">
                <h3>📁 제출 파일 (${data.파일목록.length}개)</h3>
                <ul>
`;

    data.파일목록.forEach(file => {
      htmlContent += `                    <li>${file.이름} <span style="color: #95a5a6;">(${file.타입})</span></li>\n`;
    });

    htmlContent += `                </ul>
            </div>

            <div class="content-section">
                <h3>📖 학습 내용</h3>
                <div class="content-body">
`;

    // 마크다운을 HTML로 변환
    const cleanedContent = 마크다운클린업(data.내용);
    const htmlBody = 마크다운을HTML로(cleanedContent);
    htmlContent += htmlBody;

    htmlContent += `
                </div>
            </div>
        </div>
`;
  });

  htmlContent += `
    </div>
</body>
</html>`;

  const htmlFileName = `digest-${dateStr}.html`;
  const existingHtml = folder.getFilesByName(htmlFileName);
  while (existingHtml.hasNext()) {
    existingHtml.next().setTrashed(true);
  }

  const htmlFile = folder.createFile(htmlFileName, htmlContent, MimeType.HTML);
  htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log(`\n파일 저장 완료:`);
  Logger.log(`  - ${fullFileName} (전체 원본 내용)`);
  Logger.log(`  - ${summaryFileName} (간단 요약)`);
  Logger.log(`  - ${jsonFileName} (JSON 데이터)`);
  Logger.log(`  - ${htmlFileName} (HTML 파일)`);

  // 웹앱 URL 생성 (digest-webapp.gs의 doGet 사용)
  // 웹앱을 배포한 후에는 아래 URL이 자동으로 생성됩니다
  Logger.log(`\n📱 카톡 공유 URL (웹앱 배포 필요):`);
  Logger.log(`배포 후: https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?date=${dateStr}`);
  Logger.log(`\n💡 웹앱 배포 방법:`);
  Logger.log(`1. Apps Script 상단 "배포" 클릭`);
  Logger.log(`2. "새 배포" 선택`);
  Logger.log(`3. 유형: "웹 앱"`);
  Logger.log(`4. 실행 계정: "나"`);
  Logger.log(`5. 액세스 권한: "모든 사용자"`);
  Logger.log(`6. 배포 클릭 → URL 복사`);
}

/**
 * 수동 실행용: 오늘의 다이제스트 테스트
 */
function AI다이제스트테스트() {
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  일일AI다이제스트생성(today);
}

/**
 * 수동 실행용: 어제의 다이제스트
 */
function 어제AI다이제스트생성() {
  일일AI다이제스트생성(); // dateStr 없으면 자동으로 어제
}

/**
 * 🔍 디버깅: 실제 폴더 구조 확인
 * 한 조원의 폴더 안에 어떤 하위 폴더들이 있는지 확인
 */
function 실제폴더구조확인() {
  // 첫 번째 조원의 폴더 ID 가져오기
  const firstMember = Object.entries(CONFIG.MEMBERS)[0];
  const memberName = firstMember[0];
  const folderIdOrArray = firstMember[1];
  const folderId = Array.isArray(folderIdOrArray) ? folderIdOrArray[0] : folderIdOrArray;

  Logger.log(`=== ${memberName} 폴더 구조 확인 ===`);
  Logger.log(`폴더 ID: ${folderId}\n`);

  try {
    const memberFolder = DriveApp.getFolderById(folderId);
    Logger.log(`📁 조원 폴더: ${memberFolder.getName()}`);
    Logger.log(`\n하위 폴더 목록:`);

    const subFolders = memberFolder.getFolders();
    let count = 0;

    while (subFolders.hasNext() && count < 20) {  // 최대 20개만 출력
      const folder = subFolders.next();
      const folderName = folder.getName();

      Logger.log(`  ${count + 1}. ${folderName}`);

      // 첫 번째 하위 폴더의 내부도 확인
      if (count === 0) {
        Logger.log(`     └─ ${folderName} 안의 하위 폴더:`);
        const subSubFolders = folder.getFolders();
        let subCount = 0;
        while (subSubFolders.hasNext() && subCount < 10) {
          const subFolder = subSubFolders.next();
          Logger.log(`        ${subCount + 1}. ${subFolder.getName()}`);
          subCount++;
        }
      }

      count++;
    }

    if (count === 0) {
      Logger.log(`  ❌ 하위 폴더가 없습니다!`);
    } else {
      Logger.log(`\n총 ${count}개의 하위 폴더가 있습니다.`);
    }

  } catch (e) {
    Logger.log(`❌ 오류 발생: ${e.message}`);
    Logger.log(e.stack);
  }
}

/**
 * 🔍 다이제스트 저장 폴더 확인
 * 다이제스트 파일이 저장되는 폴더의 이름과 URL을 출력
 */
function AI저장폴더확인() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
    const folderName = folder.getName();
    const folderUrl = folder.getUrl();

    Logger.log(`=== 다이제스트 저장 위치 ===`);
    Logger.log(`폴더명: ${folderName}`);
    Logger.log(`폴더 URL: ${folderUrl}`);
    Logger.log(`폴더 ID: ${CONFIG.JSON_FOLDER_ID}`);

    Logger.log(`\n최근 생성된 다이제스트 파일:`);
    const files = folder.getFiles();
    const digestFiles = [];

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      if (fileName.startsWith('full-content-') || fileName.startsWith('summary-') || fileName.startsWith('digest-')) {
        digestFiles.push({
          name: fileName,
          date: file.getLastUpdated(),
          url: file.getUrl()
        });
      }
    }

    // 날짜 순으로 정렬
    digestFiles.sort((a, b) => b.date - a.date);

    // 최근 10개만 출력
    digestFiles.slice(0, 10).forEach((file, index) => {
      Logger.log(`  ${index + 1}. ${file.name}`);
      Logger.log(`     수정일: ${Utilities.formatDate(file.date, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')}`);
      Logger.log(`     URL: ${file.url}`);
    });

    if (digestFiles.length === 0) {
      Logger.log(`  ❌ 다이제스트 파일이 없습니다.`);
    }

  } catch (e) {
    Logger.log(`❌ 오류 발생: ${e.message}`);
    Logger.log(e.stack);
  }
}
