# 📁 JSON 파일 ID 설정 완벽 가이드

## 🎯 목표
HTML 파일이 Google Drive에서 JSON 파일을 불러올 수 있도록 파일 ID를 설정합니다.

---

## 1단계: 파일 ID 찾기 (Apps Script)

### A. Apps Script 열기
```
1. Google 스프레드시트 열기
2. 확장 프로그램 > Apps Script
```

### B. 파일 ID 확인 함수 실행

**새 함수 추가:**
```javascript
// Apps Script에 이 함수를 추가하세요
function JSON파일ID확인() {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  // 현재 연월 계산
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const yearMonth = year + '-' + String(month).padStart(2, '0');

  Logger.log('=== JSON 파일 ID 목록 ===');
  Logger.log('');

  // 1. 일간 출석 파일
  const attendanceFileName = `attendance_summary_${yearMonth}.json`;
  const attendanceFiles = folder.getFilesByName(attendanceFileName);

  if (attendanceFiles.hasNext()) {
    const file = attendanceFiles.next();
    const fileId = file.getId();
    const url = `https://drive.google.com/uc?export=download&id=${fileId}`;

    Logger.log('📄 일간 출석 파일:');
    Logger.log('  파일명: ' + attendanceFileName);
    Logger.log('  파일 ID: ' + fileId);
    Logger.log('  URL: ' + url);
    Logger.log('');
  } else {
    Logger.log('❌ 일간 출석 파일을 찾을 수 없습니다: ' + attendanceFileName);
    Logger.log('');
  }

  // 2. 주간 집계 파일
  const weeklyFileName = `weekly_summary_${yearMonth}.json`;
  const weeklyFiles = folder.getFilesByName(weeklyFileName);

  if (weeklyFiles.hasNext()) {
    const file = weeklyFiles.next();
    const fileId = file.getId();
    const url = `https://drive.google.com/uc?export=download&id=${fileId}`;

    Logger.log('📊 주간 집계 파일:');
    Logger.log('  파일명: ' + weeklyFileName);
    Logger.log('  파일 ID: ' + fileId);
    Logger.log('  URL: ' + url);
    Logger.log('');
  } else {
    Logger.log('❌ 주간 집계 파일을 찾을 수 없습니다: ' + weeklyFileName);
    Logger.log('  → 이번달주간집계() 함수를 먼저 실행하세요!');
    Logger.log('');
  }

  Logger.log('======================');
  Logger.log('위의 파일 ID들을 복사해서 HTML에 붙여넣으세요!');
}
```

### C. 함수 실행 및 ID 복사

```
1. 함수 선택: JSON파일ID확인
2. 실행 (▶️) 클릭
3. 실행 로그 확인
4. 다음과 같은 결과가 나옵니다:
```

**예시 출력:**
```
=== JSON 파일 ID 목록 ===

📄 일간 출석 파일:
  파일명: attendance_summary_2025-11.json
  파일 ID: 1a2b3c4d5e6f7g8h9i0j
  URL: https://drive.google.com/uc?export=download&id=1a2b3c4d5e6f7g8h9i0j

📊 주간 집계 파일:
  파일명: weekly_summary_2025-11.json
  파일 ID: 9z8y7x6w5v4u3t2s1r0q
  URL: https://drive.google.com/uc?export=download&id=9z8y7x6w5v4u3t2s1r0q

======================
위의 파일 ID들을 복사해서 HTML에 붙여넣으세요!
```

**📝 메모장에 복사:**
```
일간 출석 ID: 1a2b3c4d5e6f7g8h9i0j
주간 집계 ID: 9z8y7x6w5v4u3t2s1r0q
```

---

## 2단계: HTML 파일 수정

### 방법 A: GitHub 웹에서 수정 (추천)

#### ① GitHub 파일 열기
```
1. https://github.com/floating535-lang/study-attendance 접속
2. index.html 클릭
3. 연필 아이콘 (✏️ Edit this file) 클릭
```

#### ② 수정할 부분 찾기

**Ctrl + F로 검색:** `JSON_FILE_IDS`

다음 코드를 찾습니다:

```javascript
// 현재 코드 (수정 전)
const baseUrl = 'YOUR_GOOGLE_DRIVE_JSON_FOLDER_URL';

// 또는 index-완성본.html을 복사했다면:
const baseUrl = 'https://drive.google.com/uc?export=download&id=';
const JSON_FILE_IDS = {
    attendance: 'YOUR_ATTENDANCE_FILE_ID',
    weekly: 'YOUR_WEEKLY_FILE_ID'
};
```

#### ③ 파일 ID 입력

**1단계에서 복사한 ID를 붙여넣습니다:**

```javascript
// 수정 후
const baseUrl = 'https://drive.google.com/uc?export=download&id=';
const JSON_FILE_IDS = {
    attendance: '1a2b3c4d5e6f7g8h9i0j',  // ← 일간 출석 파일 ID
    weekly: '9z8y7x6w5v4u3t2s1r0q'       // ← 주간 집계 파일 ID
};
```

**⚠️ 주의사항:**
- 작은따옴표 `'` 안에 ID만 입력
- 쉼표 `,` 빼먹지 않기
- URL 전체가 아니라 **ID만** 입력

#### ④ 저장하기
```
1. 아래로 스크롤
2. Commit message: "Update JSON file IDs"
3. "Commit changes" 버튼 클릭
```

---

### 방법 B: 로컬에서 수정 후 푸시

#### ① 파일 열기
```bash
cd /path/to/study-attendance
code index.html  # 또는 다른 에디터
```

#### ② 같은 방식으로 수정
```javascript
const JSON_FILE_IDS = {
    attendance: '1a2b3c4d5e6f7g8h9i0j',
    weekly: '9z8y7x6w5v4u3t2s1r0q'
};
```

#### ③ Git 커밋 & 푸시
```bash
git add index.html
git commit -m "Update JSON file IDs"
git push origin main
```

---

## 3단계: 작동 확인

### A. GitHub Pages 접속
```
https://floating535-lang.github.io/study-attendance/
```

### B. 개발자 도구로 확인

**F12 키를 누르고:**

```
1. Console 탭 클릭
2. 에러 메시지 확인
```

**✅ 성공 시:**
```
(출력 없음 또는)
데이터 로드 완료
```

**❌ 실패 시:**
```
Failed to load JSON: https://drive.google.com/uc?export=download&id=undefined
→ 파일 ID가 제대로 입력되지 않음

또는

Access denied
→ 파일 공유 권한 문제
```

### C. 화면 확인

**보여야 할 것:**
- ✅ 상단: 일간 출석표
- ✅ 하단: 주간 출석 집계 테이블
- ✅ 주차별 인증/결석 데이터

---

## 🔧 문제 해결

### 문제 1: "파일을 찾을 수 없습니다"

**증상:**
```
실행 로그:
❌ 주간 집계 파일을 찾을 수 없습니다: weekly_summary_2025-11.json
```

**해결:**
```javascript
// Apps Script에서 실행
이번달주간집계()
```
→ 주간 집계 JSON 파일을 먼저 생성해야 합니다.

---

### 문제 2: "Access denied" 에러

**증상:**
```
Console:
Failed to load JSON: Access to XMLHttpRequest has been blocked
```

**원인:** 파일 공유 권한이 비공개로 설정됨

**해결:**
```javascript
// Apps Script에서 실행
function 파일공유설정() {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();

    // 링크가 있는 모든 사용자에게 읽기 권한 부여
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('공유 설정 완료: ' + file.getName());
  }
}
```

---

### 문제 3: 매달 ID가 바뀌나요?

**예, 바뀝니다!**

현재 시스템은 매달 새로운 JSON 파일을 생성합니다:
- `attendance_summary_2025-11.json`
- `attendance_summary_2025-12.json`
- ...

#### 해결책 A: 매달 수동으로 변경
```
매달 1일에:
1. JSON파일ID확인() 실행
2. 새 파일 ID 복사
3. HTML 수정
```

#### 해결책 B: 자동화 (권장)

**HTML을 다음과 같이 수정:**

```javascript
// 수정 전 (매달 바꿔야 함)
const dataUrl = `${baseUrl}${JSON_FILE_IDS.attendance}`;
const weeklyUrl = `${baseUrl}${JSON_FILE_IDS.weekly}`;

// 수정 후 (자동으로 현재 달 파일 선택)
async function loadAttendanceData() {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const yearMonth = `${year}-${month}`;

    // 폴더에서 파일 검색하는 방식으로 변경
    const folderUrl = 'https://drive.google.com/drive/folders/YOUR_FOLDER_ID';

    // 또는 파일명으로 직접 접근
    const attendanceFileName = `attendance_summary_${yearMonth}.json`;
    const weeklyFileName = `weekly_summary_${yearMonth}.json`;

    // ... (파일명 기반으로 로드)
}
```

하지만 이 방법은 **폴더 ID**를 사용해야 합니다.

#### 해결책 C: 고정 파일명 사용 (가장 간단)

**Apps Script에서 "최신" 파일을 고정 이름으로 복사:**

```javascript
function 최신JSON복사() {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const yearMonth = year + '-' + String(month).padStart(2, '0');

  // 1. 이번 달 파일 찾기
  const attendanceFile = folder.getFilesByName(`attendance_summary_${yearMonth}.json`).next();
  const weeklyFile = folder.getFilesByName(`weekly_summary_${yearMonth}.json`).next();

  // 2. "latest" 이름으로 복사 (기존 파일 덮어쓰기)
  const latestAttendance = folder.getFilesByName('attendance_summary_latest.json');
  if (latestAttendance.hasNext()) {
    latestAttendance.next().setTrashed(true);
  }
  attendanceFile.makeCopy('attendance_summary_latest.json', folder);

  const latestWeekly = folder.getFilesByName('weekly_summary_latest.json');
  if (latestWeekly.hasNext()) {
    latestWeekly.next().setTrashed(true);
  }
  weeklyFile.makeCopy('weekly_summary_latest.json', folder);

  Logger.log('최신 JSON 파일 복사 완료');
}
```

**그러면 HTML에서:**
```javascript
const JSON_FILE_IDS = {
    attendance: '고정된_latest_파일_ID',  // 한 번만 설정하면 끝!
    weekly: '고정된_latest_파일_ID'
};
```

→ 매달 변경할 필요 없음!

---

## 📋 체크리스트

설정 완료 확인:

- [ ] Apps Script에서 `JSON파일ID확인()` 실행
- [ ] 두 개의 파일 ID 복사 (일간, 주간)
- [ ] HTML에서 `JSON_FILE_IDS` 찾기
- [ ] 파일 ID 붙여넣기
- [ ] 작은따옴표, 쉼표 확인
- [ ] GitHub에 커밋
- [ ] https://floating535-lang.github.io/study-attendance/ 접속
- [ ] F12 → Console 에러 확인
- [ ] 주간 집계 테이블 확인

---

## 🎯 최종 코드 예시

**Apps Script (JSON파일ID확인 함수 추가):**
```javascript
function JSON파일ID확인() {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const yearMonth = year + '-' + String(month).padStart(2, '0');

  Logger.log('=== JSON 파일 ID 목록 ===');

  const attendanceFile = folder.getFilesByName(`attendance_summary_${yearMonth}.json`);
  if (attendanceFile.hasNext()) {
    Logger.log('일간 출석 ID: ' + attendanceFile.next().getId());
  }

  const weeklyFile = folder.getFilesByName(`weekly_summary_${yearMonth}.json`);
  if (weeklyFile.hasNext()) {
    Logger.log('주간 집계 ID: ' + weeklyFile.next().getId());
  }
}
```

**HTML (index.html 수정 부분):**
```javascript
// 설정 영역
const baseUrl = 'https://drive.google.com/uc?export=download&id=';
const JSON_FILE_IDS = {
    attendance: '여기에_일간_출석_파일_ID',
    weekly: '여기에_주간_집계_파일_ID'
};

// 데이터 로드 함수
async function loadAttendanceData() {
    try {
        // 일간 출석 데이터
        const attendanceUrl = `${baseUrl}${JSON_FILE_IDS.attendance}`;
        const attendanceResponse = await fetch(attendanceUrl);
        const attendanceData = await attendanceResponse.json();

        // 주간 집계 데이터
        const weeklyUrl = `${baseUrl}${JSON_FILE_IDS.weekly}`;
        const weeklyResponse = await fetch(weeklyUrl);
        const weeklyData = await weeklyResponse.json();

        // 렌더링
        renderAttendanceTable(attendanceData);
        renderWeeklySummary(weeklyData);

    } catch (error) {
        console.error('데이터 로드 실패:', error);
    }
}
```

---

## 💡 팁

1. **북마크 추천:**
   - Apps Script: 빠른 접근을 위해 북마크
   - JSON_FILE_IDS 코드 라인: 에디터에서 북마크

2. **매달 1일 루틴:**
   ```
   1. 이번달주간집계() 실행
   2. JSON파일ID확인() 실행
   3. HTML 파일 ID 업데이트
   ```

3. **자동화 트리거 설정:**
   - `이번달주간집계()` → 매일 자동 실행
   - `최신JSON복사()` → 매일 자동 실행
   - HTML은 변경 불필요!

---

이제 JSON 파일 ID 설정이 명확해졌나요? 추가로 궁금한 부분이 있으면 말씀해주세요! 😊
