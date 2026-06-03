// ════════════════════════════════════════════════════════════════
// ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
// 1. Откройте ваш Google Sheets
// 2. Меню: Расширения → Apps Script
// 3. Удалите весь существующий код
// 4. Вставьте этот файл целиком
// 5. Измените SHEET_NAME на точное название вашего листа
// 6. Нажмите "Сохранить" (Ctrl+S)
// 7. Нажмите "Развернуть" → "Новое развёртывание"
// 8. Тип: "Веб-приложение"
// 9. Выполнять как: "Я" (ваш аккаунт)
// 10. Доступ: "Все" (Anyone)
// 11. Нажмите "Развернуть" → скопируйте URL
// 12. Вставьте URL в дашборды РОПа и маркетолога
// ════════════════════════════════════════════════════════════════

// ⚠️ ИЗМЕНИТЕ на точное название вашего листа (вкладка внизу)
const SHEET_NAME = "РНП";

// Ячейки с данными — не менять
const CELLS = {
  revenue: "AQ33",  // Чистая выручка
  leads:   "AQ4",   // Количество лидов
  roas:    "AQ9",   // ROAS
};

function doGet(e) {
  // CORS-заголовок — разрешает запросы с GitHub Pages
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Content-Type": "application/json",
  };

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    const revenue = parseNumber(sheet.getRange(CELLS.revenue).getValue());
    const leads   = parseNumber(sheet.getRange(CELLS.leads).getValue());
    const roas    = parseNumber(sheet.getRange(CELLS.roas).getValue());

    const data = JSON.stringify({
      ok:        true,
      updatedAt: new Date().toLocaleString("ru-RU", { timeZone: "Asia/Almaty" }),
      revenue:   revenue,
      leads:     leads,
      roas:      roas,
    });

    return ContentService
      .createTextOutput(data)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Preflight для CORS
function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

// Очищает значение — убирает пробелы, валюту, заменяет запятые
function parseNumber(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return val;
  const cleaned = String(val)
    .replace(/\s/g, "")
    .replace(/[^\d.,\-]/g, "")
    .replace(",", ".");
  return parseFloat(cleaned) || 0;
}
