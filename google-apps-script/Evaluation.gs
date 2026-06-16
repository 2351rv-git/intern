// Google Apps Script - 실습평가 시스템 (분기별 학교/학생 + 스프레드시트 연동)
// Google Apps Script 에디터에 이 코드를 붙여넣고 웹앱으로 배포하세요.
//
// ★ 스프레드시트 설정:
// 1. 새 구글 스프레드시트를 만들고 ID를 아래에 입력
// 2. "Students" 시트를 만들고 아래 형식으로 입력:
//    A열: 분기 (예: 2025-하계)
//    B열: 학교 (예: 부산대학교)  
//    C열: 학생 (예: 김철수)
// 3. 1행은 헤더로 사용 (분기 / 학교 / 학생)

const EVAL_SPREADSHEET_ID = '1yLXiMbbsPPP8dkxwyFOfZJ3xsENRDNvRjV3beRTdix8';

function doGet(e) { return handleEvalRequest(e); }
function doPost(e) { return handleEvalRequest(e); }

function handleEvalRequest(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch (action) {
      case 'saveEvaluation':
        result = saveEvaluation(JSON.parse(e.postData.contents));
        break;
      case 'getEvaluations':
        result = getEvaluations();
        break;
      case 'getStudents':
        result = getStudents();
        break;
      case 'saveStudents':
        result = saveStudents(JSON.parse(e.postData.contents));
        break;
      case 'getPassword':
        result = getPassword();
        break;
      case 'savePassword':
        result = savePassword(JSON.parse(e.postData.contents));
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (error) {
    result = { error: error.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 학생 명단 관리 (Students 시트) =====

/**
 * Students 시트에서 분기별 학교/학생 데이터 읽기
 * 시트 형식 (시트명: 2026 등 연도별):
 *   A=분기, B=기간, C=학교, D=학생, E=성별, F=공문, G=회신서, H=주간평가(O/X), I=연락처
 *   (1행 헤더)
 */
function getStudents() {
  const ss = SpreadsheetApp.openById(EVAL_SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const semesters = {};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // Evaluations, WebApp 시트는 건너뛰기
    if (sheetName === 'Evaluations' || sheetName === 'WebApp') return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    let currentSemester = '';
    let currentSchool = '';
    let currentPeriod = '';

    for (let i = 1; i < data.length; i++) {
      // A열: 분기 (빈칸이면 이전 값 유지)
      const semRaw = String(data[i][0]).trim();
      if (semRaw) currentSemester = semRaw;

      // B열: 기간 (빈칸이면 이전 값 유지)
      const periodRaw = String(data[i][1] || '').trim();
      if (periodRaw) currentPeriod = periodRaw;
      
      // C열: 학교 (빈칸이면 이전 값 유지)
      // D열에 값이 있고 C열에도 값이 있으면 → C열은 학교
      // D열에 값이 있고 C열이 빈칸이면 → 이전 학교 유지
      // D열이 빈칸이고 C열에 값이 있으면 → C열은 학교 (학생 미배정 행)
      const schoolRaw = String(data[i][2]).trim();
      if (schoolRaw) currentSchool = schoolRaw;

      // D열: 학생
      const student = String(data[i][3]).trim();

      // H열: 주간평가 (O/X)
      const weeklyRaw = String(data[i][7] || '').trim().toUpperCase();
      const weekly = (weeklyRaw === 'O' || weeklyRaw === 'YES' || weeklyRaw === 'Y');

      if (!currentSemester || !currentSchool) continue;
      if (!semesters[currentSemester]) semesters[currentSemester] = {};
      if (!semesters[currentSemester][currentSchool]) {
        semesters[currentSemester][currentSchool] = { students: [], weekly: false };
      }

      // 주간평가 여부
      if (weekly) semesters[currentSemester][currentSchool].weekly = true;

      // 학생 추가
      if (student && !semesters[currentSemester][currentSchool].students.includes(student)) {
        semesters[currentSemester][currentSchool].students.push(student);
      }

      // 학생별 상세정보 저장 (성별, 기간)
      if (student) {
        if (!semesters[currentSemester][currentSchool].details) {
          semesters[currentSemester][currentSchool].details = {};
        }
        const gender = String(data[i][4] || '').trim();
        semesters[currentSemester][currentSchool].details[student] = {
          gender: gender,
          period: currentPeriod
        };
      }
    }
  });

  return { semesters };
}

/**
 * 웹앱에서 Students 시트로 데이터 저장
 * 기존 시트에 새 시트('WebApp')로 저장하여 원본 보존
 */
function saveStudents(data) {
  const ss = SpreadsheetApp.openById(EVAL_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('WebApp');
  if (!sheet) sheet = ss.insertSheet('WebApp');
  sheet.clear();

  const headers = ['분기', '학교', '학생', '주간평가'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const rows = [];
  const semesters = data.semesters || {};
  Object.keys(semesters).sort().forEach(semester => {
    const schools = semesters[semester];
    Object.keys(schools).sort().forEach(school => {
      const schoolData = schools[school];
      const students = schoolData.students || [];
      const weeklyMark = schoolData.weekly ? 'O' : 'X';
      if (students.length === 0) {
        rows.push([semester, school, '', weeklyMark]);
      } else {
        students.sort().forEach(student => {
          rows.push([semester, school, student, weeklyMark]);
        });
      }
    });
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  return { success: true, message: `${rows.length}건 저장 완료` };
}

// ===== 평가 결과 관리 (Evaluations 시트) =====

function saveEvaluation(data) {
  const ss = SpreadsheetApp.openById(EVAL_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Evaluations');

  if (!sheet) {
    sheet = ss.insertSheet('Evaluations');
    const headers = [
      '타임스탬프', '분기', '평가자', '학교', '학생',
      '가치와방향', '실습태도', '규칙준수', '업무수행', '대인관계',
      '총점', '주간평가여부', '주간평가코멘트'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  const row = [
    data.timestamp || new Date().toISOString(),
    data.semester || '',
    data.evaluator || '',
    data.school || '',
    data.student || '',
    data.scores.section1 || 0,
    data.scores.section2 || 0,
    data.scores.section3 || 0,
    data.scores.section4 || 0,
    data.scores.section5 || 0,
    data.totalScore || 0,
    data.weeklyApplies || '',
    data.comment || ''
  ];

  sheet.appendRow(row);
  return { success: true, message: '평가가 저장되었습니다.' };
}

function getEvaluations() {
  const ss = SpreadsheetApp.openById(EVAL_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Evaluations');
  if (!sheet) return { evaluations: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { evaluations: [] };

  const evaluations = [];
  for (let i = 1; i < data.length; i++) {
    evaluations.push({
      timestamp: data[i][0],
      semester: data[i][1],
      evaluator: data[i][2],
      school: data[i][3],
      student: data[i][4],
      scores: {
        section1: data[i][5],
        section2: data[i][6],
        section3: data[i][7],
        section4: data[i][8],
        section5: data[i][9]
      },
      totalScore: data[i][10],
      weeklyApplies: data[i][11],
      comment: data[i][12]
    });
  }
  return { evaluations };
}

function tryParseJSON(str) {
  try { return JSON.parse(str); }
  catch (e) { return []; }
}

// ===== 관리자 비밀번호 관리 (Settings 시트) =====

function getPassword() {
  const ss = SpreadsheetApp.openById(EVAL_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) return { password: '' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'adminPassword') {
      return { password: String(data[i][1]).trim() };
    }
  }
  return { password: '' };
}

function savePassword(data) {
  const ss = SpreadsheetApp.openById(EVAL_SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }

  const values = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === 'adminPassword') {
      sheet.getRange(i + 1, 2).setValue(data.password);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow(['adminPassword', data.password]);
  }

  return { success: true, message: '비밀번호가 변경되었습니다.' };
}
