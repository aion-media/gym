/**
 * =========================================================================
 * 🏋️ GYM COACH API — Синхронизация тренировок с вкладкой "New GYM2"
 * =========================================================================
 */

const CONFIG = {
  SHEET_NAME: "New GYM2",        // Имя новой вкладки (New GYM2 / NewGym2)
  HEADER_COLOR: "#2ea043",      // Зеленый цвет для ячейки даты завершенной тренировки
  HEADER_TEXT_COLOR: "#ffffff"
};

function doGet(e) {
  try {
    const sheet = getWorkSheet();
    const blockNum = (e && e.parameter && e.parameter.block) ? parseInt(e.parameter.block) : 1;
    const data = parseWorkoutBlock(sheet, blockNum);
    return createJsonResponse({ success: true, sheetName: sheet.getName(), block: blockNum, data: data });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString(), stack: err.stack });
  }
}

function doPost(e) {
  try {
    const sheet = getWorkSheet();
    let payload = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : (e.parameter || {});
    const saveResult = saveWorkoutToSheet(sheet, payload);
    return createJsonResponse({ success: true, message: "Тренировка успешно сохранена!", result: saveResult });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString(), stack: err.stack });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Надежный поиск листа "New GYM2"
 */
function getWorkSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Прямой поиск
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet) return sheet;

  // 2. Поиск без учета пробелов и регистра (newgym2)
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const clean = sheets[i].getName().replace(/\s+/g, '').toLowerCase();
    if (clean === "newgym2") return sheets[i];
  }

  // 3. Если не найдена NewGym2, ищем NewGym
  for (let i = 0; i < sheets.length; i++) {
    const clean = sheets[i].getName().replace(/\s+/g, '').toLowerCase();
    if (clean === "newgym") return sheets[i];
  }

  // 4. Первый лист
  return sheets[0];
}

/**
 * Парсинг упражнений, двух последних столбцов и дат
 */
function parseWorkoutBlock(sheet, blockNum) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  let block1Row = -1;
  let block2Row = -1;

  for (let r = 0; r < values.length; r++) {
    const cellVal = String(values[r][0] || "").trim().toLowerCase();
    if (cellVal.indexOf("блок 1") !== -1 || cellVal.indexOf("блок1") !== -1) block1Row = r;
    else if (cellVal.indexOf("блок 2") !== -1 || cellVal.indexOf("блок2") !== -1) block2Row = r;
  }
  if (block1Row === -1) block1Row = 0;

  const startRow = (blockNum === 2 && block2Row !== -1) ? block2Row : block1Row;
  const endRow = (blockNum === 1 && block2Row !== -1) ? block2Row : values.length;

  // Ищем строку с надписью "Дата" в столбце A в рамках этого блока
  let dateRowIdx = -1;
  for (let r = startRow; r < Math.min(startRow + 5, endRow); r++) {
    const cellA = String(values[r][0] || "").trim().toLowerCase();
    if (cellA.indexOf("дата") !== -1) {
      dateRowIdx = r;
      break;
    }
  }
  if (dateRowIdx === -1) dateRowIdx = startRow + 1; // по умолчанию следующая строка

  // Находим все столбцы с данными для этого блока (столбцы с индекса 1 / col B и далее)
  const colsWithData = [];
  for (let c = 1; c < lastCol; c++) {
    let hasData = false;
    for (let r = dateRowIdx + 1; r < endRow; r++) {
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
    prevColIdx = colsWithData[0];
    targetColIdx = colsWithData[0] + 1;
  } else {
    prevColIdx = 1;
    targetColIdx = 2;
  }

  // Получаем дату прошлой тренировки из ячейки даты
  let prevDateVal = (prevColIdx >= 0 && dateRowIdx >= 0) ? values[dateRowIdx][prevColIdx] : "";
  let prevDateStr = formatDateValue(prevDateVal) || "Прошлая";

  // Дата на сегодня
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+3", "dd.MM.yyyy");

  // Парсим упражнения блока
  const exercises = [];
  let currentExercise = null;

  for (let r = dateRowIdx + 1; r < endRow; r++) {
    const rawName = String(values[r][0] || "").trim();

    if (rawName && !rawName.toLowerCase().startsWith("блок") && !rawName.toLowerCase().startsWith("дата")) {
      if (currentExercise && currentExercise.sets.length > 0) {
        exercises.push(currentExercise);
      }
      currentExercise = { name: rawName.replace(/\n/g, " "), sets: [] };
    }

    if (currentExercise) {
      const prevVal = prevColIdx >= 0 ? String(values[r][prevColIdx] || "").trim() : "";
      const targetVal = targetColIdx < lastCol ? String(values[r][targetColIdx] || "").trim() : "";

      if (prevVal || targetVal) {
        currentExercise.sets.push({
          row: r + 1, // 1-based для Google Таблицы
          setNum: currentExercise.sets.length + 1,
          prev: prevVal || targetVal || "—",
          target: targetVal || prevVal || ""
        });
      }
    }
  }

  if (currentExercise && currentExercise.sets.length > 0) {
    exercises.push(currentExercise);
  }

  return {
    block: blockNum,
    dateRow: dateRowIdx + 1,    // 1-based строка для записи даты
    prevCol: prevColIdx + 1,    // 1-based колонка прошлой тренировки
    targetCol: targetColIdx + 1,// 1-based колонка сегодняшней тренировки
    prevDate: prevDateStr,      // Дата прошлой тренировки (например, "30.09.2026")
    targetDate: todayStr,       // Текущая дата (например, "03.11.2026")
    exercises: exercises
  };
}

/**
 * Преобразование даты из ячейки в строку dd.MM.yyyy
 */
function formatDateValue(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || "GMT+3", "dd.MM.yyyy");
  }
  const s = String(val).trim().replace(/\[.*?\]/g, "").trim();
  return s;
}

/**
 * Сохранение тренировки с проставлением ДАТЫ в нужное место
 */
function saveWorkoutToSheet(sheet, payload) {
  let targetCol = parseInt(payload.targetCol || 0);
  const dateRow = parseInt(payload.dateRow || 2);
  const dateStr = payload.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+3", "dd.MM.yyyy");
  const results = payload.results || [];

  if (targetCol <= 0) targetCol = sheet.getLastColumn() + 1;
  if (targetCol > sheet.getMaxColumns()) sheet.insertColumnAfter(sheet.getMaxColumns());

  // 1. Проставляем ДАТУ тренировки в ячейку даты (например, C2 или C44) со стильным зеленым бейджем
  const dateCell = sheet.getRange(dateRow, targetCol);
  dateCell.setValue(dateStr);
  dateCell.setBackground(CONFIG.HEADER_COLOR);
  dateCell.setFontColor(CONFIG.HEADER_TEXT_COLOR);
  dateCell.setFontWeight("bold");
  dateCell.setHorizontalAlignment("center");

  // 2. Проставляем фактические значения каждого выполненного подхода
  let savedCount = 0;
  results.forEach(item => {
    if (item.row && item.value && item.completed) {
      const cell = sheet.getRange(item.row, targetCol);
      cell.setValue(item.value);
      cell.setHorizontalAlignment("center");
      savedCount++;
    }
  });

  return {
    savedSets: savedCount,
    column: targetCol,
    dateRow: dateRow,
    date: dateStr
  };
}

function testWorkoutData() {
  const sheet = getWorkSheet();
  Logger.log("Лист: " + sheet.getName());
  const b1 = parseWorkoutBlock(sheet, 1);
  Logger.log("Блок 1: Прошлая дата=" + b1.prevDate + ", Сегодня=" + b1.targetDate + ", Упражнений=" + b1.exercises.length);
  const b2 = parseWorkoutBlock(sheet, 2);
  Logger.log("Блок 2: Прошлая дата=" + b2.prevDate + ", Сегодня=" + b2.targetDate + ", Упражнений=" + b2.exercises.length);
}
