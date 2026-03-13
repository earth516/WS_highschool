/**
 * @OnlyCurrentDoc
 */

function getSs() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('신의경애 스티커판')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSheets() {
  const ss = getSs();
  
  const SHEET_GRADES = ["1학년", "2학년", "3학년"];
  const SHEET_LOG = "스티커 배부기록";
  const SHEET_TEACHERS = "교사 명단";

  // 1~3학년 시트 자동 생성
  SHEET_GRADES.forEach(grade => {
    let sheet = ss.getSheetByName(grade);
    if (!sheet) {
      sheet = ss.insertSheet(grade);
      sheet.appendRow(["학번", "이름", "총 스티커 개수"]);
    }
  });

  let sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog) {
    sheetLog = ss.insertSheet(SHEET_LOG);
    // [업데이트] 헤더에 '사유' 추가
    sheetLog.appendRow(["일시", "학번", "이름", "배부 개수", "배부 교사", "사유"]);
  } else {
    // 기존 시트가 있다면 헤더 확인 후 없으면 추가 (선택 사항, 수동 추가 권장)
    const header = sheetLog.getRange(1, 6).getValue();
    if (header !== "사유") {
      sheetLog.getRange(1, 6).setValue("사유");
      sheetLog.getRange(1, 1).setValue("일시"); // 날짜 -> 일시로 명칭 변경
    }
  }

  let sheetTeachers = ss.getSheetByName(SHEET_TEACHERS);
  if (!sheetTeachers) {
    sheetTeachers = ss.insertSheet(SHEET_TEACHERS);
    sheetTeachers.appendRow(["교사 이름", "비밀번호", "잔여량(자동)", "사용량(자동)", "1차 배부량", "2차 배부량", "3차 배부량"]);
  }
}

function getStudentList() {
  const ss = getSs();
  let allStudents = [];
  const SHEET_GRADES = ["1학년", "2학년", "3학년"];

  // 3개 학년 시트를 모두 돌면서 학생 명단을 하나로 합침
  SHEET_GRADES.forEach(grade => {
    const sheet = ss.getSheetByName(grade);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        const students = data.slice(1).map(row => ({ id: row[0], name: row[1] }));
        allStudents = allStudents.concat(students);
      }
    }
  });
  return allStudents;
}

function loginTeacher(name, password) {
  const ss = getSs();
  const sheet = ss.getSheetByName("교사 명단");
  if (!sheet) return { success: false, message: "교사 명단 시트가 없습니다." };
  
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) == name && String(data[i][1]) == password) {
      // E열(인덱스 4)부터 끝까지 N차 배부량 총합 계산
      let totalQuota = 0;
      for (let c = 4; c < data[i].length; c++) {
        totalQuota += Number(data[i][c] || 0);
      }
      
      // D열(인덱스 3)의 누적 사용량
      let used = Number(data[i][3] || 0);
      let remaining = totalQuota - used;
      
      return { success: true, limit: remaining };
    }
  }
  return { success: false, message: "이름 또는 비밀번호가 일치하지 않습니다." };
}

function checkStudentStatus(studentId) {
  const ss = getSs();
  const SHEET_GRADES = ["1학년", "2학년", "3학년"];

  // 3개 학년 시트를 순서대로 검색하여 학생을 찾음
  for (const grade of SHEET_GRADES) {
    const sheet = ss.getSheetByName(grade);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(studentId)) {
        return { 
          success: true, 
          name: data[i][1], 
          total: Number(data[i][2] || 0) 
        };
      }
    }
  }
  return { success: false, message: "해당 학번의 학생을 찾을 수 없습니다." };
}

// [업데이트] reason 매개변수 추가
function giveSticker(teacherName, password, studentId, count, reason) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { success: false, message: "잠시 후 다시 시도해주세요." };
  }

  const ss = getSs();
  const sheetTeachers = ss.getSheetByName("교사 명단");
  const sheetLog = ss.getSheetByName("스티커 배부기록");
  
  const teacherData = sheetTeachers.getDataRange().getValues();
  
  let teacherRowIndex = -1;
  let currentTotalQuota = 0;
  let currentUsed = 0;
  let currentLimit = 0;

  for (let i = 1; i < teacherData.length; i++) {
    if (String(teacherData[i][0]) == teacherName && String(teacherData[i][1]) == password) {
      teacherRowIndex = i + 1;
      
      // E열(인덱스 4)부터 끝까지 N차 배부량 총합 계산
      for (let c = 4; c < teacherData[i].length; c++) {
        currentTotalQuota += Number(teacherData[i][c] || 0);
      }
      currentUsed = Number(teacherData[i][3] || 0);
      currentLimit = currentTotalQuota - currentUsed;
      break;
    }
  }

  if (teacherRowIndex === -1) {
    lock.releaseLock();
    return { success: false, message: "인증 실패: 교사 정보가 올바르지 않습니다." };
  }

  if (currentLimit < count) {
    lock.releaseLock();
    return { success: false, message: `잔여량이 부족합니다. (현재: ${currentLimit}개)` };
  }

  const SHEET_GRADES = ["1학년", "2학년", "3학년"];
  let targetSheet = null;
  let studentRowIndex = -1;
  let studentName = "";
  let currentTotal = 0;

  // 3개 학년 시트를 순서대로 검색
  for (const grade of SHEET_GRADES) {
    const sheet = ss.getSheetByName(grade);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(studentId)) {
        targetSheet = sheet;
        studentRowIndex = i + 1;
        studentName = data[i][1];
        currentTotal = Number(data[i][2] || 0);
        break;
      }
    }
    if (targetSheet) break; // 학생을 찾으면 즉시 검색 종료
  }

  if (!targetSheet) {
    lock.releaseLock();
    return { success: false, message: "학생 정보를 찾을 수 없습니다." };
  }

  // 교사 잔여량 및 사용량 업데이트 로직 (C열, D열)
  let newUsed = currentUsed + Number(count);
  let newRemaining = currentTotalQuota - newUsed;
  sheetTeachers.getRange(teacherRowIndex, 3).setValue(newRemaining); // C열: 자동계산된 잔여량 
  sheetTeachers.getRange(teacherRowIndex, 4).setValue(newUsed);      // D열: 누적 사용량
  
  // 학생이 소속된 학년 시트에 개수 업데이트
  targetSheet.getRange(studentRowIndex, 3).setValue(currentTotal + Number(count));
  
  const now = new Date();
  // 시트 구조: 일시 | 학번 | 이름 | 배부 개수 | 배부 교사 | 사유
  sheetLog.appendRow([now.toLocaleString(), studentId, studentName, count, teacherName, reason]);

  SpreadsheetApp.flush();
  lock.releaseLock();

  return { success: true, remaining: currentLimit - count, studentName: studentName };
}