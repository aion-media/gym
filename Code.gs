/**
 * =========================================================================
 * 🏋️ GYM COACH API — Скрипт синхронизации тренировок с Google Таблицей
 * Вкладка: "New Gym"
 * 
 * Логика:
 * - Предпоследний столбец = Прошлая тренировка (только для чтения)
 * - Последний столбец = План на сегодня (то, что заполнил тренер)
 * =========================================================================
 */

const CONFIG = {
  SHEET_NAME: "New Gym",
  HEADER_COLOR: "#2ea043",
  HEADER_TEXT_COLOR: "#ffffff"
};

function doGet(e) {
  try {
    const sheet = getWorkSheet();
    const blockNum = (e && e.parameter && e.parameter.block) ? parseInt(e.parameter.block) : 1;
    const data = parseWorkoutBlock(sheet, blockNum);
    return createJsonResponse({ success: true, block: blockNum, data: data });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const sheet = getWorkSheet();
    let payload = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : (e.parameter || {});
    const saveResult = saveWorkoutToSheet(sheet, payload);
    return createJsonResponse({ success: true, message: "Тренировка успешно сохранена!", result: saveResult });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getWorkSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim().toLowerCase() === CONFIG.SHEET_NAME.toLowerCase()) return sheets[i];
    }
    sheet = sheets[0];
  }
  return sheet;
}

function parseWorkoutBlock(sheet, blockNum) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  let block1Row = -1, block2Row = -1;
  for (let r = 0; r < values.length; r++) {
    const cellVal = String(values[r][0] || "").trim().toLowerCase();
    if (cellVal.indexOf("блок 1") !== -1 || cellVal.indexOf("блок1") !== -1) block1Row = r;
    else if (cellVal.indexOf("блок 2") !== -1 || cellVal.indexOf("блок2") !== -1) block2Row = r;
  }
  if (block1Row === -1) block1Row = 0;

  const startRow = (blockNum === 2 && block2Row !== -1) ? block2Row : block1Row;
  const endRow = (blockNum === 1 && block2Row !== -1) ? block2Row : values.length;
  const dateRowIdx = startRow + 1;

  // Находим все столбцы с данными для этого блока
  const colsWithData = [];
  for (let c = 2; c < lastCol; c++) {
    let hasData = false;
    for (let r = startRow + 2; r < endRow; r++) {
      if (String(values[r][c] || "").trim() !== "") {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      colsWithData.push(c);
    }
  }

  // Два последних столбца: предпоследний = прошлая, последний = сегодня
  let prevColIdx = -1;
  let targetColIdx = -1;

  if (colsWithData.length >= 2) {
    prevColIdx = colsWithData[colsWithData.length - 2];
    targetColIdx = colsWithData[colsWithData.length - 1];
  } else if (colsWithData.length === 1) {
    prevColIdx = 1; // Столбец B (база)
    targetColIdx = colsWithData[0];
  } else {
    prevColIdx = 1;
    targetColIdx = 1;
  }

  const exercises = [];
  let currentExercise = null;

  for (let r = startRow + 2; r < endRow; r++) {
    const rawName = String(values[r][0] || "").trim();
    if (rawName && !rawName.toLowerCase().startsWith("блок")) {
      if (currentExercise && currentExercise.sets.length > 0) {
        exercises.push(currentExercise);
      }
      currentExercise = { name: rawName, sets: [] };
    }

    if (currentExercise) {
      const prevVal = prevColIdx >= 0 ? String(values[r][prevColIdx] || "").trim() : "";
      const targetVal = targetColIdx >= 0 ? String(values[r][targetColIdx] || "").trim() : "";
      const baseVal = String(values[r][1] || "").trim();

      if (prevVal || targetVal || baseVal) {
        currentExercise.sets.push({
          row: r + 1,
          setNum: currentExercise.sets.length + 1,
          prev: prevVal || baseVal || "—",
          target: targetVal || prevVal || baseVal || ""
        });
      }
    }
  }

  if (currentExercise && currentExercise.sets.length > 0) {
    exercises.push(currentExercise);
  }

  const prevHeader = prevColIdx >= 0 ? String(values[dateRowIdx][prevColIdx] || "").replace(/\[.*?\]/g, "").trim() : "Прошлая";
  const targetHeader = targetColIdx >= 0 ? String(values[dateRowIdx][targetColIdx] || "").replace(/\[.*?\]/g, "").trim() : "Сегодня";

  return {
    block: blockNum,
    dateRow: dateRowIdx + 1,
    prevCol: prevColIdx + 1,
    targetCol: targetColIdx + 1,
    prevDate: prevHeader || "Прошлая",
    targetDate: targetHeader || "Сегодня",
    exercises: exercises
  };
}

function saveWorkoutToSheet(sheet, payload) {
  let targetCol = parseInt(payload.targetCol || 0);
  const dateRow = parseInt(payload.dateRow || 10);
  const dateStr = payload.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+3", "dd.MM");
  const results = payload.results || [];

  if (targetCol <= 0) targetCol = sheet.getLastColumn() + 1;
  if (targetCol > sheet.getMaxColumns()) sheet.insertColumnAfter(sheet.getMaxColumns());

  const headerCell = sheet.getRange(dateRow, targetCol);
  headerCell.setValue(dateStr + " [Готово]");
  headerCell.setBackground(CONFIG.HEADER_COLOR);
  headerCell.setFontColor(CONFIG.HEADER_TEXT_COLOR);
  headerCell.setFontWeight("bold");
  headerCell.setHorizontalAlignment("center");

  let savedCount = 0;
  results.forEach(item => {
    if (item.row && item.value && item.completed) {
      const cell = sheet.getRange(item.row, targetCol);
      cell.setValue(item.value);
      cell.setHorizontalAlignment("center");
      savedCount++;
    }
  });

  return { savedSets: savedCount, column: targetCol, date: dateStr };
}

function testWorkoutData() {
  const sheet = getWorkSheet();
  const dataB1 = parseWorkoutBlock(sheet, 1);
  Logger.log("Блок 1 прочитан: " + dataB1.exercises.length + " упражнений.");
  if (dataB1.exercises.length > 0) {
    Logger.log("Упражнение 1: " + dataB1.exercises[0].name);
    Logger.log("Подходы: " + JSON.stringify(dataB1.exercises[0].sets));
  }
}
