# 📚 스터디 지식 공유 활성화 솔루션 가이드

## 🎯 목표
기존 **출석체크 자동화**를 유지하면서, **카카오톡에서 실시간 지식 교류**를 활성화합니다.

---

## 📊 현재 시스템 분석

### ✅ 장점 (유지할 것)
- Google Drive + Apps Script 기반 자동 출석체크
- 오프/장기오프/결석 자동 판정
- 마감시간 자동 처리
- HTML 출석표 시각화

### ❌ 문제점
- 공부 내용이 Drive에만 저장됨
- 카톡에서 지식 교류 사라짐
- 서로의 공부를 보려면 Drive 폴더를 일일이 확인해야 함

---

## 💡 솔루션: "스터디 일일 다이제스트" 시스템

### 핵심 아이디어
```
📁 Drive 파일 업로드 (기존 방식 유지)
    ↓
🤖 자동으로 내용 추출 & 요약 (Apps Script)
    ↓
🎨 예쁜 카드형 웹페이지 생성
    ↓
💬 카카오톡에 간편하게 공유 (원클릭)
```

### 3가지 공유 방식 (선택 가능)

#### 🅰 방식 A: 자동 알림 링크 (가장 추천 ⭐)
- **작동 방식**: 매일 저녁 8시 자동으로 요약 + 링크 생성
- **카톡 공유**: 방장/담당자가 하루 한 번 메시지 복사 → 카톡에 붙여넣기
- **장점**: 한 사람만 수고, 모든 조원이 링크로 쉽게 확인
- **예시 메시지**:
  ```
  📚 오늘의 스터디 다이제스트 (2025-11-22)

  ✅ 출석: 7명
  🏖️ 오프: 2명

  🌟 하이라이트:
  센트룸님이 5개 파일 업로드
  길님의 본초학 정리가 상세함

  🔗 자세히 보기: https://...
  ```

#### 🅱 방식 B: 개인별 자동 공유
- **작동 방식**: 각자 웹페이지에서 "내 공부 카톡 공유" 버튼 클릭
- **카톡 공유**: 자동으로 포맷된 텍스트가 클립보드에 복사됨 → 카톡에 붙여넣기
- **장점**: 개인별로 원하는 시간에 공유 가능
- **단점**: 각자 수동으로 공유해야 함

#### 🅲 방식 C: 전체 요약 텍스트
- **작동 방식**: 전체 조원의 공부 내용을 하나의 텍스트로 요약
- **카톡 공유**: 한 사람이 하루 한 번 전체 요약 복사 → 카톡에 붙여넣기
- **장점**: 카톡 내에서 바로 읽을 수 있음
- **단점**: 텍스트가 길어질 수 있음

---

## 🛠️ 구현 단계

### 1단계: Apps Script 확장 (콘텐츠 수집 & 가공)

#### 📄 `apps script code.gs`에 추가할 코드

```javascript
// ==================== 일일 다이제스트 기능 ====================

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

  // 방장에게 이메일 발송 (선택 사항)
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

  // 다이제스트 페이지 URL (실제 배포 시 변경)
  const digestUrl = 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec?page=digest';
  message += `\n🔗 자세히 보기: ${digestUrl}\n`;
  message += `\n💡 Tip: 링크를 클릭하면 조원들의 공부 내용을 예쁘게 볼 수 있어요!`;

  return message;
}

/**
 * 수동 실행용: 오늘의 다이제스트 메시지 로그 출력
 */
function 오늘의다이제스트메시지확인() {
  일일다이제스트생성();
}
```

---

### 2단계: HTML 다이제스트 페이지 생성

#### 📄 새 파일: `daily-digest.html`

Apps Script 웹앱으로 배포할 HTML 파일을 생성합니다.

**방법 1: Apps Script 내에 HTML 파일 추가**
1. Google Sheets → 확장 프로그램 → Apps Script
2. 좌측 메뉴에서 `+` 버튼 → HTML 파일 추가
3. 파일명: `digest-page`
4. 아래 코드 붙여넣기

```html
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📚 오늘의 스터디 다이제스트</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Malgun Gothic', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            margin: 0;
            padding: 20px;
            min-height: 100vh;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
        }
        .header {
            text-align: center;
            color: white;
            margin-bottom: 30px;
        }
        .header h1 {
            font-size: 32px;
            margin: 0 0 10px 0;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }
        .header .date {
            font-size: 18px;
            opacity: 0.9;
        }

        .stats {
            background: rgba(255,255,255,0.95);
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 30px;
            display: flex;
            justify-content: space-around;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .stat-item {
            text-align: center;
        }
        .stat-number {
            font-size: 32px;
            font-weight: bold;
            color: #667eea;
        }
        .stat-label {
            font-size: 14px;
            color: #666;
            margin-top: 5px;
        }

        .digest-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 20px;
            margin-bottom: 100px;
        }

        .member-card {
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: all 0.3s ease;
            position: relative;
        }
        .member-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 12px rgba(0,0,0,0.2);
        }

        .member-card.absent {
            opacity: 0.6;
            border: 2px dashed #ccc;
        }

        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #f0f0f0;
        }

        .member-name {
            font-size: 20px;
            font-weight: bold;
            color: #667eea;
        }

        .status-badge {
            font-size: 24px;
        }

        .card-title {
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 10px;
            color: #333;
            line-height: 1.4;
        }

        .card-thumbnail {
            width: 100%;
            height: 150px;
            object-fit: cover;
            border-radius: 8px;
            margin-bottom: 15px;
        }

        .card-summary {
            font-size: 14px;
            color: #666;
            line-height: 1.6;
            margin-bottom: 15px;
            max-height: 120px;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .card-files {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 15px;
        }

        .file-badge {
            background: #f0f0f0;
            padding: 5px 10px;
            border-radius: 15px;
            font-size: 12px;
            color: #666;
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }
        .file-badge a {
            color: inherit;
            text-decoration: none;
        }
        .file-badge:hover {
            background: #e0e0e0;
        }

        .copy-btn {
            width: 100%;
            padding: 10px;
            background: #667eea;
            color: white;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.2s;
        }
        .copy-btn:hover {
            background: #5568d3;
        }
        .copy-btn.copied {
            background: #10b981;
        }

        .share-all-btn {
            position: fixed;
            bottom: 30px;
            right: 30px;
            padding: 15px 30px;
            background: #f59e0b;
            color: white;
            border: none;
            border-radius: 50px;
            font-size: 16px;
            font-weight: bold;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            cursor: pointer;
            z-index: 100;
            transition: all 0.2s;
        }
        .share-all-btn:hover {
            background: #d97706;
            transform: scale(1.05);
        }

        .loading {
            text-align: center;
            padding: 50px;
            color: white;
            font-size: 18px;
        }

        .empty-message {
            text-align: center;
            padding: 40px;
            color: #999;
            font-style: italic;
        }

        @media (max-width: 768px) {
            .digest-grid {
                grid-template-columns: 1fr;
            }
            .share-all-btn {
                bottom: 20px;
                right: 20px;
                padding: 12px 24px;
                font-size: 14px;
            }
            .stats {
                flex-direction: column;
                gap: 15px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📚 오늘의 스터디 다이제스트</h1>
            <div class="date" id="digestDate">로딩 중...</div>
        </div>

        <div class="stats" id="statsContainer" style="display: none;">
            <div class="stat-item">
                <div class="stat-number" id="attendCount">0</div>
                <div class="stat-label">✅ 출석</div>
            </div>
            <div class="stat-item">
                <div class="stat-number" id="offCount">0</div>
                <div class="stat-label">🏖️ 오프</div>
            </div>
            <div class="stat-item">
                <div class="stat-number" id="absentCount">0</div>
                <div class="stat-label">❌ 결석</div>
            </div>
        </div>

        <div class="loading" id="loadingIndicator">
            ⏳ 데이터 로드 중...
        </div>

        <div class="digest-grid" id="digestGrid" style="display: none;">
            <!-- 카드는 JS로 동적 생성 -->
        </div>

        <button class="share-all-btn" id="shareAllBtn" onclick="copyAllToKakao()" style="display: none;">
            💬 전체 내용 카톡 공유
        </button>
    </div>

    <script>
        let digestData = null;

        // 페이지 로드 시 데이터 가져오기
        window.addEventListener('DOMContentLoaded', async () => {
            await loadDigestData();
        });

        async function loadDigestData() {
            try {
                // URL 파라미터에서 날짜 가져오기 (없으면 오늘)
                const urlParams = new URLSearchParams(window.location.search);
                const dateParam = urlParams.get('date');
                const targetDate = dateParam || formatDate(new Date());

                // JSON 파일 URL 생성
                const jsonUrl = `https://drive.google.com/uc?export=download&id=JSON_FILE_ID_${targetDate}`;

                // 개발용: Apps Script 웹앱 엔드포인트에서 가져오기
                const response = await fetch('https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec?action=getDigest&date=' + targetDate);

                if (!response.ok) {
                    throw new Error('데이터 로드 실패');
                }

                digestData = await response.json();

                renderDigest(digestData);

            } catch (error) {
                console.error('데이터 로드 오류:', error);
                document.getElementById('loadingIndicator').innerHTML =
                    '❌ 데이터를 불러올 수 없습니다.<br><small>' + error.message + '</small>';
            }
        }

        function renderDigest(data) {
            document.getElementById('loadingIndicator').style.display = 'none';
            document.getElementById('digestGrid').style.display = 'grid';
            document.getElementById('shareAllBtn').style.display = 'block';
            document.getElementById('statsContainer').style.display = 'flex';

            // 날짜 표시
            const dateObj = new Date(data.date);
            document.getElementById('digestDate').textContent =
                `${dateObj.getFullYear()}년 ${dateObj.getMonth() + 1}월 ${dateObj.getDate()}일`;

            // 통계 계산
            let attendCount = 0;
            let offCount = 0;
            let absentCount = 0;

            const grid = document.getElementById('digestGrid');
            grid.innerHTML = '';

            for (const [name, content] of Object.entries(data.members)) {
                // 통계 집계
                if (content.상태 === '출석') attendCount++;
                else if (content.상태 === '오프' || content.상태 === '장기오프') offCount++;
                else if (content.상태 === '결석') absentCount++;

                // 카드 생성
                const card = createMemberCard(name, content);
                grid.appendChild(card);
            }

            // 통계 표시
            document.getElementById('attendCount').textContent = attendCount;
            document.getElementById('offCount').textContent = offCount;
            document.getElementById('absentCount').textContent = absentCount;
        }

        function createMemberCard(name, content) {
            const card = document.createElement('div');
            card.className = 'member-card';

            if (content.상태 !== '출석') {
                card.classList.add('absent');
            }

            let html = `
                <div class="card-header">
                    <div class="member-name">${escapeHtml(name)}</div>
                    <div class="status-badge">${getStatusEmoji(content.상태)}</div>
                </div>
            `;

            if (content.상태 === '출석') {
                if (content.제목) {
                    html += `<div class="card-title">${escapeHtml(content.제목)}</div>`;
                }

                if (content.썸네일) {
                    html += `<img class="card-thumbnail" src="${content.썸네일}" alt="학습 이미지" loading="lazy">`;
                }

                if (content.요약) {
                    html += `<div class="card-summary">${escapeHtml(content.요약)}</div>`;
                }

                if (content.파일목록 && content.파일목록.length > 0) {
                    html += '<div class="card-files">';
                    content.파일목록.forEach(file => {
                        const icon = getFileIcon(file.타입);
                        html += `
                            <div class="file-badge">
                                <a href="${file.링크}" target="_blank" rel="noopener">
                                    ${icon} ${escapeHtml(file.이름)}
                                </a>
                            </div>
                        `;
                    });
                    html += '</div>';
                }

                html += `
                    <button class="copy-btn" onclick="copyMemberContent('${escapeHtml(name)}')">
                        📋 카톡에 공유하기
                    </button>
                `;
            } else {
                html += `<div class="empty-message">오늘은 ${content.상태}</div>`;
            }

            card.innerHTML = html;
            return card;
        }

        function getStatusEmoji(status) {
            switch(status) {
                case '출석': return '✅';
                case '오프': return '🏖️';
                case '장기오프': return '🏝️';
                case '결석': return '❌';
                case '제출대기': return '⏰';
                default: return '❓';
            }
        }

        function getFileIcon(type) {
            switch(type) {
                case 'Markdown': return '📝';
                case 'PDF': return '📕';
                case 'Image': return '🖼️';
                default: return '📄';
            }
        }

        function copyMemberContent(memberName) {
            const content = digestData.members[memberName];

            let text = `📚 ${memberName}님의 오늘 공부\n\n`;
            if (content.제목) text += `📖 ${content.제목}\n\n`;
            if (content.요약) text += `${content.요약}\n\n`;
            if (content.파일목록 && content.파일목록.length > 0) {
                text += `📎 파일: ${content.파일목록.map(f => f.이름).join(', ')}\n`;
            }
            if (content.폴더링크) {
                text += `\n🔗 자세히 보기: ${content.폴더링크}`;
            }

            copyToClipboard(text, event.target);
        }

        function copyAllToKakao() {
            let text = `📚 오늘의 스터디 다이제스트 (${digestData.date})\n\n`;

            let attendList = [];
            for (const [name, content] of Object.entries(digestData.members)) {
                if (content.상태 === '출석') {
                    let summary = `✅ ${name}`;
                    if (content.제목) summary += `\n   ${content.제목}`;
                    if (content.요약) {
                        const shortSummary = content.요약.substring(0, 80).trim();
                        summary += `\n   ${shortSummary}${content.요약.length > 80 ? '...' : ''}`;
                    }
                    attendList.push(summary);
                }
            }

            text += attendList.join('\n\n');
            text += `\n\n━━━━━━━━━━━━━━━━`;
            text += `\n📊 출석: ${document.getElementById('attendCount').textContent}명`;
            text += ` | 오프: ${document.getElementById('offCount').textContent}명`;
            text += ` | 결석: ${document.getElementById('absentCount').textContent}명`;
            text += `\n\n🔗 자세히 보기: ${window.location.href}`;

            copyToClipboard(text, event.target);
        }

        function copyToClipboard(text, button) {
            navigator.clipboard.writeText(text).then(() => {
                const originalText = button.textContent;
                button.textContent = '✅ 복사 완료!';
                button.classList.add('copied');

                setTimeout(() => {
                    button.textContent = originalText;
                    button.classList.remove('copied');
                }, 2000);
            }).catch(err => {
                alert('복사 실패: ' + err.message);
            });
        }

        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        function formatDate(date) {
            const y = date.getFullYear();
            const m = String(date.getMonth() + 1).padStart(2, '0');
            const d = String(date.getDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
        }
    </script>
</body>
</html>
```

---

### 3단계: Apps Script 웹앱 배포

#### `Code.gs`에 웹앱 엔드포인트 추가

```javascript
/**
 * 웹앱 진입점 - HTML 페이지 제공
 */
function doGet(e) {
  const page = e.parameter.page || 'digest';
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
```

#### 배포 단계:
1. Apps Script 편집기에서 **배포 > 새 배포**
2. 유형: **웹 앱**
3. 설명: "스터디 다이제스트 v1"
4. 실행 계정: **나**
5. 액세스 권한: **모든 사용자** (링크를 아는 사람만)
6. **배포** 클릭
7. **웹 앱 URL 복사** → 이것이 다이제스트 페이지 주소

---

### 4단계: 트리거 설정

Apps Script 편집기에서:

1. **시계 아이콘 (트리거)** 클릭
2. **트리거 추가**:
   - 실행할 함수: `일일다이제스트생성`
   - 배포 선택: Head
   - 이벤트 소스: 시간 기반
   - 시간 기반 트리거 유형: 일 단위 타이머
   - 시간대 선택: 오후 8시 ~ 오후 9시
3. **저장**

---

## 📱 카카오톡 공유 방법

### 방법 A: 자동 알림 링크 (추천)

1. **매일 저녁 8시**: Apps Script가 자동으로 다이제스트 생성
2. **방장/담당자**: 다이제스트 페이지 URL 방문
3. **"전체 내용 카톡 공유" 버튼** 클릭 → 자동으로 복사됨
4. **카카오톡**에 붙여넣기

**예시 메시지**:
```
📚 오늘의 스터디 다이제스트 (2025-11-22)

✅ 센트룸
   한방병리학 - 담음의 병리
   담음의 형성 원인과 병리 기전에 대해 공부. 비위의 운화 실조가 수습 대사 장애를 일으켜...

✅ 길
   본초학 - 해표약
   해표약의 분류와 각 약재의 특성 비교. 발산풍한약과 발산풍열약의 차이점, 주요 약재의...

━━━━━━━━━━━━━━━━
📊 출석: 7명 | 오프: 2명 | 결석: 0명

🔗 자세히 보기: https://script.google.com/...
```

### 방법 B: 개인별 공유

각 조원이:
1. 다이제스트 페이지 방문
2. 자기 카드에서 **"카톡에 공유하기"** 버튼 클릭
3. 카카오톡에 붙여넣기

### 방법 C: 자동 이메일 알림 (선택)

`일일다이제스트생성()` 함수에 이메일 발송 코드 추가:
```javascript
GmailApp.sendEmail(
  '방장이메일@example.com',
  '[자동] 오늘의 스터디 다이제스트',
  message
);
```

---

## 🎯 기대 효과

### ✅ 문제 해결
- ✓ 출석체크 자동화 유지
- ✓ 카톡에서 실시간 지식 교류 활성화
- ✓ 공부 내용 자동 요약 및 시각화
- ✓ 참여자 부담 최소화 (기존 업로드 방식 그대로)
- ✓ 카톡 외 플랫폼 불필요

### 📈 추가 효과
- 조원들의 공부 자극 증가
- 서로의 학습 주제 파악 가능
- 질문 및 피드백 활성화
- 학습 내용 아카이빙 (웹페이지로 보관)

---

## 🔧 고급 기능 (선택 사항)

### 1. 이미지 OCR 추가
```javascript
function 이미지OCR처리(imageFile) {
  try {
    const resource = {
      title: imageFile.getName(),
      mimeType: MimeType.GOOGLE_DOCS
    };

    const doc = Drive.Files.copy(resource, imageFile.getId(), {
      ocr: true,
      ocrLanguage: 'ko'
    });

    const docFile = DocumentApp.openById(doc.id);
    const text = docFile.getBody().getText();

    DriveApp.getFileById(doc.id).setTrashed(true);

    return text;
  } catch (e) {
    return '';
  }
}
```

### 2. AI 요약 (OpenAI API 연동)
```javascript
function AI요약생성(fullText) {
  const apiKey = 'YOUR_OPENAI_API_KEY';
  const url = 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: 'gpt-3.5-turbo',
    messages: [{
      role: 'user',
      content: `다음 공부 내용을 2-3문장으로 요약해주세요:\n\n${fullText}`
    }],
    max_tokens: 150
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    return json.choices[0].message.content.trim();
  } catch (e) {
    return fullText.substring(0, 200) + '...';
  }
}
```

### 3. 주간 다이제스트 생성
```javascript
function 주간다이제스트생성() {
  const today = new Date();
  const weekAgo = new Date(today);
  weekAgo.setDate(today.getDate() - 7);

  // 지난 7일간의 다이제스트 합치기
  // ...
}
```

---

## 📞 지원 및 문의

문제가 발생하면:
1. Apps Script 로그 확인 (Ctrl+Enter)
2. 트리거 실행 기록 확인
3. JSON 파일 생성 확인

---

## 🎉 마무리

이 시스템을 통해:
- 기존 **자동화의 편리함**을 유지하면서
- **카카오톡에서 지식 공유**를 활성화하고
- **참여자 부담 없이** 운영할 수 있습니다!

**핵심**: 매일 저녁 자동으로 예쁜 다이제스트가 생성되고, 한 사람이 링크만 카톡에 공유하면 끝!
