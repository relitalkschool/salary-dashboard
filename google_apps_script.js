// ════════════════════════════════════════════════════════════════
// ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
// 1. Откройте ваш Google Sheets
// 2. Меню: Расширения → Apps Script
// 3. Удалите весь существующий код
// 4. Вставьте этот файл целиком
// 5. Нажмите "Сохранить" (Ctrl+S)
// 6. Нажмите "Развернуть" → "Новое развёртывание"
// 7. Тип: "Веб-приложение"
// 8. Выполнять как: "Я" (ваш аккаунт)
// 9. Доступ: "Все" (чтобы калькуляторы могли читать)
// 10. Нажмите "Развернуть" → скопируйте URL веб-приложения
// 11. Вставьте этот URL в калькуляторы РОПа и маркетолога
// ════════════════════════════════════════════════════════════════

// Название листа с вашим РНП отчётом — измените если нужно
const SHEET_NAME = "РНП";

// Ячейки с данными
const CELLS = {
  revenue: "AQ33",   // Чистая выручка
  leads:   "AQ4",    // Общее количество лидов
  roas:    "AQ9",    // ROAS
};

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME)
                  || ss.getSheets()[0]; // fallback на первый лист

    const revenue = parseNumber(sheet.getRange(CELLS.revenue).getValue());
    const leads   = parseNumber(sheet.getRange(CELLS.leads).getValue());
    const roas    = parseNumber(sheet.getRange(CELLS.roas).getValue());

    const data = {
      ok:        true,
      updatedAt: new Date().toLocaleString("ru-RU", { timeZone: "Asia/Almaty" }),
      revenue:   revenue,
      leads:     leads,
      roas:      roas,
    };

    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Очищает значение от пробелов, валют, заменяет запятые на точки
function parseNumber(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return val;
  const cleaned = String(val)
    .replace(/\s/g, "")
    .replace(/[^\d.,\-]/g, "")
    .replace(",", ".");
  return parseFloat(cleaned) || 0;
}
