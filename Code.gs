/**
 * 🌍 에코 챌린지 — Google Apps Script (Code.gs)
 * ================================================
 * [사용 방법]
 * 1. Apps Script 에디터에서 이 파일을 Code.gs에 붙여넣기
 * 2. 새 파일(HTML) 추가 → 이름: index → index.html에 프론트엔드 붙여넣기
 * 3. 배포 → 새 배포 → 웹 앱
 *    - 다음 사용자로 실행: 나(본인)
 *    - 액세스 권한: 모든 사용자
 * 4. 배포 URL로 학생들이 바로 접속 가능!
 */

const SHEET_NAME   = "에코기록";
const SUMMARY_NAME = "학생별요약";


// ─── HTML 페이지 반환 ────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  // action 없으면 → 웹 페이지 반환
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('🌍 에코 챌린지')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // action 있으면 → 데이터 API
  try {
    const ss = SpreadsheetApp.openById("1o1Ng06UJTIzyvmn3AlCmHNUKlgNUMgN5gX2QP747S-Q");

    if (action === "load") {
      const sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) return jsonResponse({ result: "success", records: [] });
      const rows    = sheet.getDataRange().getValues();
      const records = rows.slice(1)
        .filter(row => row[0] !== "")
        .map(row => ({
          date:     row[0],
          time:     row[1],
          name:     row[2],
          category: row[3],
          label:    row[4],
          points:   row[5],
          emoji:    row[6]
        }));
      return jsonResponse({ result: "success", records });
    }

    if (action === "summary") {
      const sheet = ss.getSheetByName(SUMMARY_NAME);
      if (!sheet) return jsonResponse({ result: "success", summary: [] });
      const rows    = sheet.getDataRange().getValues();
      const summary = rows.slice(1)
        .filter(row => row[0] !== "")
        .map(row => ({
          name:  row[0],
          total: row[1],
          count: row[2],
          level: row[3],
          last:  row[4]
        }));
      return jsonResponse({ result: "success", summary });
    }

    return jsonResponse({ result: "error", message: "알 수 없는 action" });

  } catch (err) {
    return jsonResponse({ result: "error", message: err.toString() });
  }
}


// ─── 기록 저장 ───────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (!data.name || !data.label || !data.points) {
      return jsonResponse({ result: "error", message: "필수 데이터 누락" });
    }

    const ss    = SpreadsheetApp.openById("1o1Ng06UJTIzyvmn3AlCmHNUKlgNUMgN5gX2QP747S-Q");
    const sheet = getOrCreateSheet(ss, SHEET_NAME, [
      "날짜", "시간", "이름", "카테고리", "실천항목", "포인트", "이모지"
    ]);

    sheet.appendRow([
      data.date,
      data.time,
      data.name,
      data.category,
      data.label,
      Number(data.points),
      data.emoji
    ]);

    updateSummary(ss, data.name, Number(data.points));

    return jsonResponse({ result: "success", message: "저장 완료" });

  } catch (err) {
    return jsonResponse({ result: "error", message: err.toString() });
  }
}


// ─── 헬퍼 함수들 ─────────────────────────────────

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headerRow = sheet.getRange(1, 1, 1, headers.length);
    headerRow.setValues([headers]);
    headerRow.setFontWeight("bold");
    headerRow.setBackground("#3b6d11");
    headerRow.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function updateSummary(ss, name, points) {
  const sheet = getOrCreateSheet(ss, SUMMARY_NAME, [
    "이름", "총포인트", "실천횟수", "레벨", "마지막실천일"
  ]);

  const data  = sheet.getDataRange().getValues();
  const today = new Date().toISOString().slice(0, 10);
  let found   = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      const newTotal = (data[i][1] || 0) + points;
      const newCount = (data[i][2] || 0) + 1;
      sheet.getRange(i + 1, 2).setValue(newTotal);
      sheet.getRange(i + 1, 3).setValue(newCount);
      sheet.getRange(i + 1, 4).setValue(getLevel(newTotal));
      sheet.getRange(i + 1, 5).setValue(today);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([name, points, 1, getLevel(points), today]);
  }
}

function getLevel(pts) {
  if (pts >= 1000) return "🌍 에코 낙원";
  if (pts >= 600)  return "🏕️ 무성한 숲";
  if (pts >= 350)  return "🌳 숲 단계";
  if (pts >= 200)  return "🌲 나무 단계";
  if (pts >= 100)  return "🌸 꽃밭 단계";
  if (pts >= 50)   return "🌷 꽃봉오리";
  if (pts >= 20)   return "🌿 새싹 단계";
  return "🌱 씨앗 단계";
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ─── 테스트 함수 ──────────────────────────────────
function testSave() {
  const ss    = SpreadsheetApp.openById("1o1Ng06UJTIzyvmn3AlCmHNUKlgNUMgN5gX2QP747S-Q");
  const sheet = getOrCreateSheet(ss, SHEET_NAME, [
    "날짜", "시간", "이름", "카테고리", "실천항목", "포인트", "이모지"
  ]);
  sheet.appendRow(["2025-06-27", "09:00", "테스트학생", "💧 물 절약", "양치할 때 컵 사용하기", 5, "🦷"]);
  updateSummary(ss, "테스트학생", 5);
  Logger.log("✅ 테스트 저장 완료!");
}
