# index.html 업데이트 가이드

## 🎯 목표

출석표 하단에 주간 통계 섹션을 추가합니다:
- 월요일 기준 주간 집계
- 진행중인 주는 결석 표시 안함
- 월요일 기준 안내문 추가

---

## 📋 수정 항목

### 1. 안내문 추가 (상단 rules-box 섹션)

**위치:** `<div class="rules-box">` 섹션 내부

**추가할 내용:**
```html
<div class="rules-box">
    <h3>📌 출석 규칙</h3>
    <ul>
        <li>매일 새벽 3시까지 공부 인증</li>
        <li>장기오프는 구글 폼으로 사전 신청</li>
        <!-- 기존 규칙들... -->
    </ul>

    <!-- 🆕 새로 추가 -->
    <h3>📅 주간 집계 기준</h3>
    <ul>
        <li><strong>주 단위:</strong> 월요일 ~ 일요일 (7일)</li>
        <li><strong>월 소속:</strong> 월요일이 속한 달의 주로 계산</li>
        <li><strong>예시:</strong> 11월 25일(월) ~ 12월 1일(일) → <strong>11월 4주차</strong></li>
        <li><strong>필요 인증:</strong> 주 4회 (장기오프 일수만큼 차감)</li>
        <li><strong>진행중인 주:</strong> 일요일이 지나지 않은 주는 결석 계산 안함</li>
    </ul>
</div>
```

---

### 2. JavaScript 데이터 로드 함수 수정

**위치:** `<script>` 섹션의 `loadAttendanceData()` 함수

**수정 내용:**

```javascript
async function loadAttendanceData() {
    const loadingDiv = document.getElementById('loading');
    const errorDiv = document.getElementById('error');

    try {
        loadingDiv.style.display = 'block';
        errorDiv.style.display = 'none';

        // 🆕 JSON 파일 URL 구성
        const baseUrl = 'YOUR_GOOGLE_DRIVE_JSON_FOLDER_URL';
        const yearMonth = `${year}-${String(month + 1).padStart(2, '0')}`;

        // 기존 출석 데이터
        const dataUrl = `${baseUrl}/attendance_summary_${yearMonth}.json`;

        // 🆕 주간 통계 데이터
        const weeklyUrl = `${baseUrl}/weekly_summary_${yearMonth}.json`;

        // 출석 데이터 로드
        const response = await fetch(dataUrl);
        if (!response.ok) {
            throw new Error(`출석 데이터 로드 실패: ${response.status}`);
        }
        const data = await response.json();

        // 🆕 주간 통계 로드
        let weeklyData = null;
        try {
            const weeklyResponse = await fetch(weeklyUrl);
            if (weeklyResponse.ok) {
                weeklyData = await weeklyResponse.json();
            }
        } catch (e) {
            console.log('주간 통계 로드 실패 (선택사항):', e);
        }

        // 전월 데이터 로드 (기존)
        let prevMonthData = null;
        // ... 기존 코드 ...

        // 테이블 렌더링
        renderAttendanceTable(data, prevMonthData, weeklyData);  // 🆕 weeklyData 추가

        loadingDiv.style.display = 'none';
    } catch (error) {
        console.error('데이터 로드 오류:', error);
        loadingDiv.style.display = 'none';
        errorDiv.style.display = 'block';
        errorDiv.textContent = `데이터 로드 실패: ${error.message}`;
    }
}
```

---

### 3. 테이블 렌더링 함수 수정

**위치:** `renderAttendanceTable()` 함수

**수정 내용:**

```javascript
function renderAttendanceTable(data, prevMonthData = null, weeklyData = null) {  // 🆕 weeklyData 파라미터 추가
    const tableBody = document.getElementById('attendanceTableBody');
    const memberSummaryDiv = document.getElementById('memberSummary');

    tableBody.innerHTML = '';
    memberSummaryDiv.innerHTML = '';

    // ... 기존 일간 출석표 렌더링 코드 ...

    // 🆕 주간 통계 렌더링
    if (weeklyData) {
        renderWeeklySummary(weeklyData);
    }
}
```

---

### 4. 주간 통계 렌더링 함수 추가 (신규)

**위치:** `<script>` 섹션 내부 (renderAttendanceTable 함수 다음)

**추가할 코드:**

```javascript
/**
 * 주간 통계 렌더링
 */
function renderWeeklySummary(weeklyData) {
    // 주간 통계 컨테이너 찾기 (없으면 생성)
    let weeklyContainer = document.getElementById('weeklySummary');

    if (!weeklyContainer) {
        // 출석표 테이블 다음에 주간 통계 섹션 추가
        const container = document.querySelector('.container');
        weeklyContainer = document.createElement('div');
        weeklyContainer.id = 'weeklySummary';
        weeklyContainer.style.marginTop = '50px';
        container.appendChild(weeklyContainer);
    }

    // HTML 생성
    let html = `
        <h2 style="text-align: center; color: #333; margin-bottom: 20px;">
            📊 주간 출석 집계 (${weeklyData.년월})
        </h2>

        <div style="background-color: #e3f2fd; border: 2px solid #2196F3; border-radius: 8px; padding: 15px; margin-bottom: 20px;">
            <h4 style="margin-top: 0; color: #1976D2;">ℹ️ ${weeklyData.안내.설명}</h4>
            <p style="margin: 5px 0;"><strong>예시:</strong> ${weeklyData.안내.예시}</p>
        </div>

        <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <thead>
                <tr>
                    <th rowspan="2" style="background-color: #2196F3; color: white;">조원</th>
    `;

    // 주차 헤더 생성
    const 주차수 = Object.values(weeklyData.조원별집계)[0]?.주차별.length || 0;
    for (let i = 1; i <= 주차수; i++) {
        html += `<th colspan="2" style="background-color: #2196F3; color: white;">${i}주차</th>`;
    }
    html += `<th rowspan="2" style="background-color: #f44336; color: white;">총결석</th>`;
    html += `</tr><tr>`;

    // 인증/결석 서브 헤더
    for (let i = 0; i < 주차수; i++) {
        html += `
            <th style="background-color: #64B5F6; color: white; font-size: 12px;">인증</th>
            <th style="background-color: #64B5F6; color: white; font-size: 12px;">결석</th>
        `;
    }
    html += `</tr></thead><tbody>`;

    // 조원별 데이터
    for (const [memberName, memberData] of Object.entries(weeklyData.조원별집계)) {
        html += `<tr>`;
        html += `<td style="font-weight: bold; text-align: left; padding-left: 15px;">${escapeHtml(memberName)}</td>`;

        // 주차별 데이터
        for (const week of memberData.주차별) {
            const 인증색 = week.인증 >= week.필요 ? '#e8f5e9' : '#ffebee';
            const 결석색 = week.결석 > 0 ? '#ffcdd2' : '#f5f5f5';

            // 진행중인 주는 "-" 표시
            const 인증표시 = week.전체장기오프 ? '🏝️' : `${week.인증}/${week.필요}`;
            const 결석표시 = week.상태 === '진행중' ? '-' : (week.전체장기오프 ? '-' : week.결석);

            html += `<td style="background-color: ${인증색};">${인증표시}</td>`;
            html += `<td style="background-color: ${결석색};">${결석표시}</td>`;
        }

        // 총결석
        const 총결석색 = memberData.총결석 >= 4 ? '#f44336' : memberData.총결석 === 3 ? '#ff9800' : '#4CAF50';
        const 총결석텍스트색 = memberData.총결석 >= 3 ? 'white' : 'black';
        html += `<td style="background-color: ${총결석색}; color: ${총결석텍스트색}; font-weight: bold;">${memberData.총결석}</td>`;

        html += `</tr>`;
    }

    html += `</tbody></table>`;

    // 범례 추가
    html += `
        <div style="margin-top: 20px; padding: 15px; background-color: #f5f5f5; border-radius: 5px;">
            <h4 style="margin-top: 0;">📖 범례</h4>
            <div style="display: flex; flex-wrap: wrap; gap: 15px;">
                <div><span style="display: inline-block; width: 20px; height: 20px; background-color: #e8f5e9; border: 1px solid #ddd; vertical-align: middle;"></span> 인증 충족</div>
                <div><span style="display: inline-block; width: 20px; height: 20px; background-color: #ffebee; border: 1px solid #ddd; vertical-align: middle;"></span> 인증 부족</div>
                <div><span style="display: inline-block; width: 20px; height: 20px; background-color: #ffcdd2; border: 1px solid #ddd; vertical-align: middle;"></span> 결석</div>
                <div><strong>-</strong> 진행중 (결석 미확정)</div>
                <div><strong>🏝️</strong> 전체 장기오프</div>
            </div>
        </div>
    `;

    weeklyContainer.innerHTML = html;
}
```

---

### 5. CSS 스타일 추가 (선택사항)

**위치:** `<style>` 섹션

**추가할 스타일:**

```css
/* 주간 통계 테이블 스타일 */
#weeklySummary table {
    font-size: 14px;
}

#weeklySummary th,
#weeklySummary td {
    border: 1px solid #ddd;
    padding: 10px 8px;
    text-align: center;
}

#weeklySummary th {
    font-weight: 600;
}

/* 반응형 */
@media (max-width: 768px) {
    #weeklySummary table {
        font-size: 12px;
    }

    #weeklySummary th,
    #weeklySummary td {
        padding: 6px 4px;
    }
}
```

---

## 🔍 JSON 파일 URL 확인

실제 JSON 파일 URL을 확인하는 방법:

1. **Apps Script에서 실행:**
   ```javascript
   function JSON폴더URL확인() {
     const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
     Logger.log('폴더 ID: ' + CONFIG.JSON_FOLDER_ID);
     Logger.log('폴더 URL: ' + folder.getUrl());

     // weekly_summary 파일 찾기
     const files = folder.getFilesByName('weekly_summary_2025-11.json');
     if (files.hasNext()) {
       const file = files.next();
       Logger.log('파일 ID: ' + file.getId());
       Logger.log('다운로드 URL: https://drive.google.com/uc?export=download&id=' + file.getId());
     }
   }
   ```

2. **URL 형식:**
   ```
   https://drive.google.com/uc?export=download&id={FILE_ID}
   ```

3. **index.html에서 사용:**
   ```javascript
   const weeklyUrl = 'https://drive.google.com/uc?export=download&id=YOUR_FILE_ID_HERE';
   ```

---

## ✅ 체크리스트

- [ ] 안내문 섹션 추가
- [ ] loadAttendanceData() 함수 수정 (주간 데이터 로드)
- [ ] renderAttendanceTable() 함수 수정 (파라미터 추가)
- [ ] renderWeeklySummary() 함수 추가
- [ ] CSS 스타일 추가
- [ ] JSON 파일 URL 확인 및 설정
- [ ] Git commit & push to floating535-lang/study-attendance
- [ ] GitHub Pages 확인

---

## 🚀 배포 순서

1. **Apps Script에서 주간 집계 실행**
   ```
   이번달주간집계() 실행
   → weekly_summary_2025-11.json 생성
   ```

2. **JSON 파일 ID 확인**
   ```
   JSON폴더URL확인() 실행
   → 파일 ID 복사
   ```

3. **index.html 수정**
   ```
   위 가이드대로 코드 추가
   JSON 파일 URL 설정
   ```

4. **Git 배포**
   ```bash
   cd study-attendance
   git add index.html
   git commit -m "Add weekly attendance summary section"
   git push
   ```

5. **확인**
   ```
   https://floating535-lang.github.io/study-attendance/
   ```

---

## 💡 팁

### 디버깅
브라우저 개발자 도구(F12)에서 확인:
```javascript
console.log('주간 데이터:', weeklyData);
```

### 테스트 데이터
주간 통계가 없을 때도 에러 없이 작동하도록 try-catch 사용

### 캐시 문제
JSON 파일이 업데이트되지 않으면:
```javascript
const weeklyUrl = `${baseUrl}/weekly_summary_${yearMonth}.json?t=${Date.now()}`;
```

---

**준비 완료하시면 실제 floating535-lang/study-attendance 리포지토리의 index.html을 수정하시면 됩니다!**
