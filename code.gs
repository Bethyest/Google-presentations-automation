var TABLE_CONFIG = {
  ROWS_PER_SLIDE:       5,
  COLUMN_WIDTHS:        [65, 73, 46, 54, 54, 52, 48, 66, 41, 40, 65, 50, 60],
  TABLE_POSITION:       { x: 3, y: 50 },
  ROW_HEIGHTS:          { header: 32, data: 28 },
  FONT_SIZES:           { header: 6, data: 6.5, dataNumber: 8 },
  BACKGROUND_IMAGE_ID:  "Your_Id"
};

var TEMPLATE_ID = "Your_id";

// ─────────────────────────────────────────────────────────────────────────────
//  ПРАВИЛА ОБЪЕДИНЕНИЯ ПО УМОЛЧАНИЮ
//
//  Каждый столбец описывается объектом:
//    type        — "all"    : правила применяются ко всем значениям
//                  "except" : правила применяются ко всем, кроме texts[]
//                  "none"   : объединение отключено полностью
//    mergeText   — направление объединения для текстовых значений:
//                  "both" | "horizontal" | "vertical" | "none"
//    mergeNumbers— направление объединения для числовых значений:
//                  "both" | "horizontal" | "vertical" | "none"
//    texts[]     — список исключений (только для type: "except")
// ─────────────────────────────────────────────────────────────────────────────

var DEFAULT_COLUMN_RULES = {
  0: { type: "all",    mergeText: "both",     mergeNumbers: "vertical" },
  1: { type: "all",    mergeText: "both",     mergeNumbers: "vertical" },
  2: { type: "none",   mergeText: "none",     mergeNumbers: "none"     },
  3: { type: "except", texts: ["Нет бюджетных мест", "Не было набора"],
                       mergeText: "both",     mergeNumbers: "vertical" },
  4: { type: "all",    mergeText: "vertical", mergeNumbers: "vertical" },
  5: { type: "all",    mergeText: "vertical", mergeNumbers: "vertical" },
  6: { type: "all",    mergeText: "vertical", mergeNumbers: "vertical" }
};

function loadColumnRules() {
  var raw = PropertiesService.getDocumentProperties().getProperty("COLUMN_RULES_JSON");
  if (!raw) return DEFAULT_COLUMN_RULES;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return DEFAULT_COLUMN_RULES;
  }
}

function saveColumnRules(rules) {
  PropertiesService.getDocumentProperties().setProperty("COLUMN_RULES_JSON", JSON.stringify(rules));
}

function resetColumnRulesToDefault() {
  PropertiesService.getDocumentProperties().deleteProperty("COLUMN_RULES_JSON");
}

function getRule(col) {
  var rules = loadColumnRules();
  return rules[col] || { type: "all", mergeText: "both", mergeNumbers: "vertical" };
}


function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("🎓 ПРЕЗЕНТАЦИЯ")
      .addItem("📊 Создать презентацию",       "createPresentation")
      .addSeparator()
      .addItem("🔄 Обновить заметки вузов",    "updateUniversityNotes")
      .addItem("⚙️ Настройки объединения",     "showMergeSidebar")
      .addSeparator()
      .addItem("✏️ Переименовать",              "renameLastPresentation")
      .addItem("🔗 Открыть",                   "openLastPresentation")
      .addToUi();
  } catch (e) {}
}

function onSelectionChange(e) {
  try {
    var range = e.range;
    if (range.getRow() === 1 && range.getSheet().getName() === "table") {
    }
  } catch (e) {}
}


function showMergeSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("MergeSettingsSidebar")
    .setTitle("⚙️ Настройки объединения")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSidebarData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("table");
  var rules = loadColumnRules();

  var colsCount  = sheet ? Math.min(sheet.getLastColumn(), TABLE_CONFIG.COLUMN_WIDTHS.length) : 0;
  var colHeaders = [];
  var colSamples = [];

  if (sheet && sheet.getLastRow() >= 1) {
    var headerRow = sheet.getRange(1, 1, 1, colsCount).getValues()[0];
    var dataRows  = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, Math.min(sheet.getLastRow() - 1, 8), colsCount).getValues()
      : [];

    for (var c = 0; c < colsCount; c++) {
      colHeaders.push(headerRow[c] ? headerRow[c].toString() : "Столбец " + (c + 1));

      var seen = {}, samples = [];
      for (var r = 0; r < dataRows.length && samples.length < 5; r++) {
        var v = dataRows[r][c];
        if (v != null && v.toString().trim() !== "" && !seen[v]) {
          seen[v] = true;
          samples.push(v.toString().trim());
        }
      }
      colSamples.push(samples);
    }
  }

  var serialized = {};
  for (var col in rules) {
    serialized[col] = rules[col];
  }

  return {
    colsCount:  colsCount,
    headers:    colHeaders,
    samples:    colSamples,
    rules:      serialized,
    defaults:   DEFAULT_COLUMN_RULES
  };
}

function applyMergeRules(rulesFromSidebar) {
  saveColumnRules(rulesFromSidebar);
  return "ok";
}

function resetMergeRules() {
  resetColumnRulesToDefault();
  return "ok";
}


function createPresentation() {
  var html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;font-size:13px">
      <h3 style="margin-top:0">📊 Создание презентации</h3>

      <div id="formStep">
        <input type="text" id="name" style="width:100%;padding:8px;margin:10px 0;box-sizing:border-box"
               placeholder="Имя презентации" value="Карта вузов">
        <button onclick="startCheck()"
                style="background:#4CAF50;color:white;padding:10px 20px;border:none;cursor:pointer">
          ✅ Создать
        </button>
        <button onclick="google.script.host.close()" style="padding:10px 20px">❌ Отмена</button>
      </div>

      <div id="issuesStep" style="display:none">
        <p style="color:#cc0000;margin:0 0 8px">⚠️ В листе "table" похоже на ошибку:</p>
        <ul id="issuesList"
            style="max-height:150px;overflow-y:auto;padding-left:18px;margin:0 0 12px;color:#444"></ul>
        <p style="font-size:12px;color:#666;margin:0 0 14px">
          Проблемные строки уже подсвечены красным, и лист открыт на первой из них.
        </p>
        <button onclick="ignoreAndCreate()"
                style="background:#ff9800;color:white;padding:10px 16px;border:none;cursor:pointer">
          ⚠️ Игнорировать и создать
        </button>
        <button onclick="google.script.host.close()" style="padding:10px 16px">⬅️ Отмена, поправлю</button>
      </div>

      <div id="statusStep" style="display:none">
        <p id="statusText" style="color:#666"></p>
      </div>
    </div>

    <script>
      var currentName = '';

      function startCheck() {
        currentName = document.getElementById('name').value || 'Карта вузов';
        showStep('statusStep');
        setStatus('🔎 Проверяю таблицу...');
        google.script.run
          .withSuccessHandler(onCheckDone)
          .withFailureHandler(onFail)
          .checkBeforeCreate(currentName);
      }

      function onCheckDone(result) {
        if (result.issues && result.issues.length > 0) {
          var list = document.getElementById('issuesList');
          list.innerHTML = '';
          result.issues.forEach(function(issue) {
            var li = document.createElement('li');
            li.textContent = issue.message;
            list.appendChild(li);
          });
          showStep('issuesStep');
        } else {
          doCreate();
        }
      }

      function ignoreAndCreate() { doCreate(); }

      function doCreate() {
        showStep('statusStep');
        setStatus('⏳ Создаю...');
        google.script.run
          .withSuccessHandler(function(url) {
            setStatus('✅ Готово!');
            setTimeout(function() { window.open(url, '_blank'); google.script.host.close(); }, 1000);
          })
          .withFailureHandler(onFail)
          .runPresentationCreationWithName(currentName);
      }

      function onFail(err)     { showStep('statusStep'); setStatus('❌ ' + err.message); }
      function setStatus(text) { document.getElementById('statusText').innerHTML = text; }

      function showStep(id) {
        ['formStep','issuesStep','statusStep'].forEach(function(s) {
          document.getElementById(s).style.display = (s === id) ? 'block' : 'none';
        });
      }
    </script>
  `).setWidth(440).setHeight(370);

  SpreadsheetApp.getUi().showModalDialog(html, 'Новая презентация');
}

function renameLastPresentation() {
  var lastId = PropertiesService.getDocumentProperties().getProperty('LAST_PRESENTATION_ID');
  if (!lastId) { SpreadsheetApp.getUi().alert('❌ Нет созданной презентации'); return; }

  var html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial">
      <h3>✏️ Переименование</h3>
      <input type="text" id="name" style="width:100%;padding:8px;margin:10px 0" placeholder="Новое имя">
      <button onclick="rename()"
              style="background:#2196F3;color:white;padding:10px 20px;border:none;cursor:pointer">
        💾 Сохранить
      </button>
      <button onclick="google.script.host.close()" style="padding:10px 20px">❌ Отмена</button>
    </div>
    <script>
      function rename() {
        var name = document.getElementById('name').value;
        if (!name) { alert('Введите имя'); return; }
        google.script.run
          .withSuccessHandler(function() { google.script.host.close(); })
          .doRenamePresentation(name);
      }
    </script>
  `).setWidth(400).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'Переименование');
}

function doRenamePresentation(newName) {
  var props = PropertiesService.getDocumentProperties();
  DriveApp.getFileById(props.getProperty('LAST_PRESENTATION_ID')).setName(newName);
  props.setProperty('LAST_PRESENTATION_NAME', newName);
  SpreadsheetApp.getUi().alert('✅ Переименовано: ' + newName);
}

function openLastPresentation() {
  var lastId = PropertiesService.getDocumentProperties().getProperty('LAST_PRESENTATION_ID');
  if (!lastId) { SpreadsheetApp.getUi().alert('❌ Нет созданной презентации'); return; }
  try {
    var url = DriveApp.getFileById(lastId).getUrl();
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(
        '<script>window.open("' + url + '","_blank");google.script.host.close();<\/script>'
      ),
      'Открытие'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Презентация не найдена');
  }
}

// НОВАЯ ФУНКЦИЯ - обновляет заметки без создания презентации
function updateUniversityNotes() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("table");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ Лист 'table' не найден");
      return;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("❌ Нет данных в таблице");
      return;
    }
    
    // Получаем данные
    var allPlain = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    var universities = getUniversitiesList(ss, allPlain);
    var headerRows = findHeaderRowIndices(allPlain);
    
    // Очищаем старые заметки
    sheet.getRange(1, 1, lastRow, 1).clearNote();
    
    // Устанавливаем новые заметки
    var count = Math.min(headerRows.length, universities.length);
    for (var k = 0; k < count; k++) {
      sheet.getRange(headerRows[k] + 1, 1).setNote("Шапка №" + (k + 1) + " — " + universities[k]);
    }
    
    SpreadsheetApp.getUi().alert("✅ Обновлено " + count + " заметок с названиями вузов");
    
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Ошибка: " + e.message);
  }
}


function isRowEmpty(row) {
  if (!row || row.length === 0) return true;
  for (var i = 0; i < row.length; i++) {
    if (row[i] && row[i].toString().trim() !== "") return false;
  }
  return true;
}

function checkBeforeCreate(presentationName) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var issues = validateAndHighlightTable(ss);
  setUniversityNotes(ss);
  return { issues: issues };
}

function validateAndHighlightTable(spreadsheet) {
  var sheet = spreadsheet.getSheetByName("table");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  sheet.getRange(2, 1, lastRow - 1, lastCol).setBackground(null);

  var data   = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var issues = [];

  var streakStart = -1;

  function recordStreak(start, end) {
    var len = end - start;
    if (len < 2) return;
    var from = start + 1, to = end;
    issues.push({
      type: "empty_rows", fromRow: from, toRow: to,
      message: "Строки " + from + "–" + to + ": " + len +
               " пустые строки подряд (между блоками должна быть только одна)"
    });
    sheet.getRange(from, 1, len, lastCol).setBackground("#f4cccc");
  }

  for (var i = 1; i < data.length; i++) {
    if (isRowEmpty(data[i])) {
      if (streakStart === -1) streakStart = i;
    } else if (streakStart !== -1) {
      recordStreak(streakStart, i);
      streakStart = -1;
    }
  }
  if (streakStart !== -1) recordStreak(streakStart, data.length);

  var inBlock = false, blockHeaderRow = -1;
  for (var i = 0; i < data.length; i++) {
    var empty = isRowEmpty(data[i]);
    if (i === 0) { inBlock = true; blockHeaderRow = 0; continue; }
    if (empty) {
      if (inBlock && blockHeaderRow >= 0) {
        issues.push({
          type: "empty_block",
          fromRow: blockHeaderRow + 1, toRow: blockHeaderRow + 1,
          message: "Строка " + (blockHeaderRow + 1) + ": шапка вуза без строк данных под ней"
        });
        sheet.getRange(blockHeaderRow + 1, 1, 1, lastCol).setBackground("#fff2cc");
      }
      inBlock = false; blockHeaderRow = -1;
    } else {
      if (!inBlock) { inBlock = true; blockHeaderRow = i; }
      else if (blockHeaderRow === i - 1 || (blockHeaderRow >= 0 && i > blockHeaderRow + 1)) {
        blockHeaderRow = -1;
      }
    }
  }
  if (inBlock && blockHeaderRow >= 0) {
    issues.push({
      type: "empty_block",
      fromRow: blockHeaderRow + 1, toRow: blockHeaderRow + 1,
      message: "Строка " + (blockHeaderRow + 1) + ": шапка вуза без строк данных под ней"
    });
    sheet.getRange(blockHeaderRow + 1, 1, 1, lastCol).setBackground("#fff2cc");
  }

  if (issues.length > 0) {
    spreadsheet.setActiveSheet(sheet);
    var f = issues[0];
    sheet.setActiveRange(
      sheet.getRange(f.fromRow, 1, Math.max(1, f.toRow - f.fromRow + 1), lastCol)
    );
  }
  return issues;
}

function findHeaderRowIndices(allPlain) {
  var headers = [0], inGap = false;
  for (var i = 1; i < allPlain.length; i++) {
    if (isRowEmpty(allPlain[i]))       { inGap = true; }
    else if (inGap)                    { headers.push(i); inGap = false; }
  }
  return headers;
}

function setUniversityNotes(spreadsheet) {
  var sheet = spreadsheet.getSheetByName("table");
  if (!sheet) return;
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  var allPlain     = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var universities = getUniversitiesList(spreadsheet, allPlain);
  var headerRows   = findHeaderRowIndices(allPlain);

  sheet.getRange(1, 1, lastRow, 1).clearNote();
  var count = Math.min(headerRows.length, universities.length);
  for (var k = 0; k < count; k++) {
    sheet.getRange(headerRows[k] + 1, 1).setNote("Шапка №" + (k + 1) + " — " + universities[k]);
  }
}

function runPresentationCreationWithName(presentationName) {
  var file   = DriveApp.getFileById(TEMPLATE_ID).makeCopy(presentationName);
  var pres   = SlidesApp.openById(file.getId());
  var presId = file.getId();
  var ss     = SpreadsheetApp.getActiveSpreadsheet();

  replaceVariables(pres, ss);

  var tableData = readTableData(ss);
  if (!tableData || !tableData.allPlain || tableData.allPlain.length === 0)
    throw new Error("Нет данных в таблице");

  var universities = getUniversitiesList(ss, tableData.allPlain);
  var uniTables    = splitDataByUniversity(tableData);
  var batchReqs    = createSlides(pres, universities, uniTables);

  pres.saveAndClose();
  if (batchReqs && batchReqs.length > 0) applyFormatting(batchReqs, presId);

  var props = PropertiesService.getDocumentProperties();
  props.setProperty("LAST_PRESENTATION_ID",   presId);
  props.setProperty("LAST_PRESENTATION_NAME", presentationName);

  return "https://docs.google.com/presentation/d/" + presId + "/edit";
}


function replaceVariables(pres, spreadsheet) {
  var sheet = spreadsheet.getSheetByName("variables");
  if (!sheet) return;
  var vars   = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  var slides = pres.getSlides();
  for (var r = 0; r < vars.length; r++) {
    var key = vars[r][0], val = vars[r][1];
    if (!key || key.toString().trim() === "") continue;
    if (!val && val !== 0) val = "";
    for (var j = 0; j < slides.length; j++)
      slides[j].replaceAllText(key.toString(), val.toString());
  }
}

function readTableData(spreadsheet) {
  var sheet = spreadsheet.getSheetByName("table");
  if (!sheet) return null;
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) return null;
  var range = sheet.getRange(1, 1, lastRow, lastCol);

  var orgFormats = range.getNumberFormats();
  range.setNumberFormat("@");
  var richGrid = range.getRichTextValues();
  range.setNumberFormats(orgFormats);

  var richGridNormal = sheet.getRange(1, 1, lastRow, lastCol).getRichTextValues();

  return {
    allRichText: richGridNormal,
    allPlain:    range.getValues(),
    allFormulas: range.getFormulas(),
    rowsCount:   lastRow,
    colsCount:   lastCol
  };
}

function extractHyperlinkUrl(formula) {
  if (!formula || typeof formula !== "string") return null;
  var m = formula.match(/^\s*=\s*HYPERLINK\s*\(\s*"((?:[^"\\]|\\.)*)"/i);
  return m && m[1] ? m[1].replace(/\\"/g, '"') : null;
}

function getUniversitiesList(spreadsheet, tableRows) {
  var sheet = spreadsheet.getSheetByName("variables");
  if (!sheet) return [];
  var vars = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();

  var universities = [];
  for (var i = 1; i < tableRows.length; i++) {
    if (isRowEmpty(tableRows[i])) universities.push("ВУЗ");
  }
  universities.push("ВУЗ");

  var uc = 0;
  for (var i = 8; i < vars.length && uc < universities.length; i++) {
    var n = vars[i] ? vars[i][1] : null;
    if (n && n.toString().trim() !== "") universities[uc++] = n.toString().trim();
  }
  for (var j = 0; j < universities.length; j++)
    if (universities[j] === "ВУЗ") universities[j] = "ВУЗ " + (j + 1);

  return universities;
}

function splitDataByUniversity(tableData) {
  var plain    = tableData.allPlain;
  var rich     = tableData.allRichText;
  var formulas = tableData.allFormulas || [];

  var result = [];
  var curRows = [], curPlain = [], curFormulas = [];
  var curHeaders = null, curHeaderFormulas = null;

  for (var i = 0; i < plain.length; i++) {
    if (i === 0) {
      curHeaders = rich[i]; curHeaderFormulas = formulas[i];
      continue;
    }
    var isEmpty = isRowEmpty(plain[i]);
    if (isEmpty) {
      if (curRows.length > 0) {
        result.push({ headers: curHeaders, headerFormulas: curHeaderFormulas,
                      rows: curRows, plain: curPlain, formulas: curFormulas });
        curRows = []; curPlain = []; curFormulas = [];
      }
      curHeaders = null; curHeaderFormulas = null;
    } else {
      if (!curHeaders) {
        curHeaders = rich[i]; curHeaderFormulas = formulas[i];
      } else {
        curRows.push(rich[i]); curPlain.push(plain[i]);
        curFormulas.push(formulas[i] || []);
      }
    }
  }
  if (curRows.length > 0)
    result.push({ headers: curHeaders, headerFormulas: curHeaderFormulas,
                  rows: curRows, plain: curPlain, formulas: curFormulas });
  return result;
}


function getColumnWidth(col) {
  return col < TABLE_CONFIG.COLUMN_WIDTHS.length ? TABLE_CONFIG.COLUMN_WIDTHS[col] : 50;
}

function getTableWidth(colsCount) {
  var total = 0;
  for (var i = 0; i < colsCount; i++) total += getColumnWidth(i);
  return total;
}

function getCenteredX(tableWidth) {
  return (720 - tableWidth) / 2;
}

function looksLikeNumber(value) {
  if (!value && value !== 0) return false;
  var str = value.toString().trim()
    .replace(/\s/g,  "")
    .replace(/,/g,   ".")
    .replace(/₽/g,   "")
    .replace(/%/g,   "");
  return str !== "" && !isNaN(str);
}

function canMergeInDirection(v1, v2, link1, link2, rule, direction) {
  if (!v1 || !v2)      return false;
  if (link1 !== link2) return false;

  var isNum1 = looksLikeNumber(v1);
  var isNum2 = looksLikeNumber(v2);
  var isNum  = isNum1 && isNum2;

  var mergeDir = isNum ? rule.mergeNumbers : rule.mergeText;

  if (!mergeDir || mergeDir === "none") return false;
  if (mergeDir !== "both" && mergeDir !== direction) return false;

  if (isNum) {
    var parse = function(s) {
      return parseFloat(s.toString().replace(/\s/g,"").replace(/,/g,".").replace(/₽/g,"").replace(/%/g,""));
    };
    return parse(v1) === parse(v2);
  }
  return v1 === v2;
}

function findGroupsInData(dataRows, dataPlain, colsCount, headers, dataFormulas, headerFormulas) {
  var hGroups = [], vGroups = [];
  var rowsCount = dataRows.length;
  if (rowsCount === 0) return { h: [], v: [] };

  function getText(row, col) {
    if (row === -1) {
      if (headers && headers[col] && typeof headers[col].getText === "function") {
        var t = headers[col].getText(); if (t && t.trim()) return t.trim();
      }
      return "";
    }
    if (dataRows[row] && dataRows[row][col] && typeof dataRows[row][col].getText === "function") {
      var t = dataRows[row][col].getText(); if (t && t.trim()) return t.trim();
    }
    if (dataPlain[row] && dataPlain[row][col] != null) {
      var t = dataPlain[row][col].toString().trim(); if (t !== "") return t;
    }
    return "";
  }

  function getFirstLink(row, col) {
    var cell    = row === -1 ? (headers && headers[col]) : (dataRows[row] && dataRows[row][col]);
    var formula = row === -1
      ? (headerFormulas && headerFormulas[col])
      : (dataFormulas && dataFormulas[row] && dataFormulas[row][col]);

    if (cell && typeof cell.getRuns === "function") {
      try {
        var runs = cell.getRuns();
        for (var i = 0; i < runs.length; i++) {
          var l = runs[i].getLinkUrl(); if (l) return l;
        }
      } catch (e) {}
    }
    return extractHyperlinkUrl(formula);
  }

  for (var row = 0; row < rowsCount; row++) {
    for (var col = 0; col < colsCount; col++) {
      var rule = getRule(col);
      if (rule.type === "none") continue;

      var t = getText(row, col);
      if (!t) continue;
      if (rule.type === "except" && rule.texts && rule.texts.indexOf(t) >= 0) continue;

      var linkT = getFirstLink(row, col);
      var end   = col;

      while (end + 1 < colsCount) {
        var nextRule = getRule(end + 1);
        if (nextRule.type === "none") break;

        var nt = getText(row, end + 1);
        if (!nt) break;
        if (nextRule.type === "except" && nextRule.texts && nextRule.texts.indexOf(nt) >= 0) break;

        if (canMergeInDirection(t, nt, linkT, getFirstLink(row, end + 1), rule, "horizontal")) {
          end++;
        } else {
          break;
        }
      }
      if (end > col) { hGroups.push({ row: row, sc: col, ec: end }); col = end; }
    }
  }

  for (var col = 0; col < colsCount; col++) {
    var rule = getRule(col);
    if (rule.type === "none") continue;

    for (var row = 0; row < rowsCount; row++) {
      var t = getText(row, col);
      if (!t) continue;
      if (rule.type === "except" && rule.texts && rule.texts.indexOf(t) >= 0) continue;

      var linkT = getFirstLink(row, col);
      var end   = row;

      while (end + 1 < rowsCount) {
        var nt = getText(end + 1, col);
        if (!nt) break;
        if (rule.type === "except" && rule.texts && rule.texts.indexOf(nt) >= 0) break;

        if (canMergeInDirection(t, nt, linkT, getFirstLink(end + 1, col), rule, "vertical")) {
          end++;
        } else {
          break;
        }
      }
      if (end > row) { vGroups.push({ sr: row, er: end, col: col }); row = end; }
    }
  }

  return { h: hGroups, v: vGroups };
}

function adjustGroupsForTable(groups) {
  var h = groups.h.map(function(g) { return { row: g.row + 1, sc: g.sc, ec: g.ec }; });
  var v = groups.v.map(function(g) { return { sr: g.sr + 1, er: g.er + 1, col: g.col }; });
  return filterOverlappingGroups(h, v);
}

function filterOverlappingGroups(hGroups, vGroups) {
  var validV = vGroups.filter(function(vg) {
    return !hGroups.some(function(hg) {
      return hg.row >= vg.sr && hg.row <= vg.er && hg.sc <= vg.col && hg.ec >= vg.col;
    });
  });
  var validH = hGroups.filter(function(hg) {
    return !validV.some(function(vg) {
      return hg.row >= vg.sr && hg.row <= vg.er && hg.sc <= vg.col && hg.ec >= vg.col;
    });
  });
  return { h: validH, v: validV };
}


function createSlides(pres, universities, uniTables) {
  var batchReqs = [];
  var SY        = TABLE_CONFIG.TABLE_POSITION.y;
  var ROWS_PER  = TABLE_CONFIG.ROWS_PER_SLIDE;

  for (var u = 0; u < Math.min(universities.length, uniTables.length); u++) {
    var name     = universities[u];
    var td       = uniTables[u];
    var rows     = td.rows, plains = td.plain || [], formulas = td.formulas || [];
    var headers  = td.headers, headerFormulas = td.headerFormulas;

    if (!rows || rows.length === 0) continue;

    var colsCount  = Math.min(headers.length, TABLE_CONFIG.COLUMN_WIDTHS.length);
    var tableWidth = getTableWidth(colsCount);
    var centeredX  = getCenteredX(tableWidth);
    var parts      = Math.ceil(rows.length / ROWS_PER);

    for (var p = 0; p < parts; p++) {
      var sr = p * ROWS_PER, er = Math.min(sr + ROWS_PER, rows.length);

      var slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slide.getBackground().setPictureFill(
        "https://drive.google.com/uc?export=view&id=" + TABLE_CONFIG.BACKGROUND_IMAGE_ID
      );

      var titleText = name + (parts > 1 ? "\n(часть " + (p + 1) + "/" + parts + ")" : "");
      var titleBox  = slide.insertTextBox(titleText, 20, 2, 665, 40);
      titleBox.getText().getTextStyle()
        .setFontFamily("Montserrat").setFontSize(12).setBold(true).setForegroundColor("#000000");
      titleBox.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      var rc    = er - sr + 1;
      var table = slide.insertTable(rc, colsCount, centeredX, SY, tableWidth, rc * 20);
      var tid   = table.getObjectId();

      var partRows     = rows.slice(sr, er);
      var partPlain    = plains.slice(sr, er);
      var partFormulas = formulas.slice(sr, er);

      var groups         = findGroupsInData(partRows, partPlain, colsCount, headers, partFormulas, headerFormulas);
      var adjustedGroups = adjustGroupsForTable(groups);

      fillTable(table, headers, partRows, partPlain, colsCount, rc,
                adjustedGroups, batchReqs, tid, partFormulas, headerFormulas);
    }
  }

  return batchReqs;
}


function fillTable(table, headers, partRows, partPlain, colsCount, rowsCount,
                   groups, batchRequests, tid, partFormulas, headerFormulas) {

  function getCellText(row, col) {
    if (row === 0) {
      if (headers && headers[col] && typeof headers[col].getText === "function")
        return (headers[col].getText() || "").trim();
      return "";
    }
    var dr = row - 1;
    if (partRows[dr] && partRows[dr][col] && typeof partRows[dr][col].getText === "function") {
      var t = partRows[dr][col].getText(); if (t && t.trim()) return t.trim();
    }
    if (partPlain[dr] && partPlain[dr][col] != null) {
      var t = partPlain[dr][col].toString().trim(); if (t !== "") return t;
    }
    return "";
  }

  function getRichCell(row, col) {
    if (row === 0) return headers && headers[col] ? headers[col] : null;
    var dr = row - 1;
    return partRows[dr] && partRows[dr][col] ? partRows[dr][col] : null;
  }

  function getCellFormula(row, col) {
    if (row === 0) return headerFormulas && headerFormulas[col];
    var dr = row - 1;
    return partFormulas && partFormulas[dr] && partFormulas[dr][col];
  }

  function isCellVisible(row, col) {
    var rule = getRule(col);
    if (rule.type === "none") return true;
    var text = getCellText(row, col);
    if (!text) return false;
    if (rule.type === "except" && rule.texts && rule.texts.indexOf(text) >= 0) return true;
    for (var i = 0; i < groups.h.length; i++) {
      var g = groups.h[i];
      if (g.row === row && col > g.sc && col <= g.ec) return false;
    }
    for (var i = 0; i < groups.v.length; i++) {
      var g = groups.v[i];
      if (g.col === col && row > g.sr && row <= g.er) return false;
    }
    return true;
  }

  function applyLinks(textRange, richCell, formula) {
    if (richCell && typeof richCell.getRuns === "function") {
      try {
        var runs = richCell.getRuns();
        var pos  = 0;
        for (var i = 0; i < runs.length; i++) {
          var runText = runs[i].getText();
          var runLink = runs[i].getLinkUrl();
          if (runLink && runText && runText.length > 0) {
            var end = pos + runText.length;
            try {
              textRange.getRange(pos, end).getTextStyle()
                .setLinkUrl(runLink).setUnderline(true).setForegroundColor("#ffffff");
            } catch (e) {}
          }
          pos += runText.length;
        }
        return;
      } catch (e) {}
    }
    var formulaLink = extractHyperlinkUrl(formula);
    if (formulaLink) {
      textRange.getTextStyle().setLinkUrl(formulaLink).setUnderline(true).setForegroundColor("#ffffff");
    }
  }

  for (var c = 0; c < colsCount; c++) {
    var cell = table.getCell(0, c);
    cell.getFill().setSolidFill("#9d8781");
    var text = getCellText(0, c);
    if (text && isCellVisible(0, c)) {
      var tr = cell.getText();
      tr.setText(text);
      tr.getTextStyle()
        .setFontFamily("Montserrat").setFontSize(TABLE_CONFIG.FONT_SIZES.header)
        .setBold(true).setForegroundColor("#ffffff");
      tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      applyLinks(tr, getRichCell(0, c), getCellFormula(0, c));
    }
  }

  for (var r = 1; r < rowsCount; r++) {
    for (var c = 0; c < colsCount; c++) {
      var cell = table.getCell(r, c);
      cell.getFill().setSolidFill("#9d8781");
      var text = getCellText(r, c);
      if (text && isCellVisible(r, c)) {
        var tr       = cell.getText();
        var fontSize = looksLikeNumber(text)
          ? TABLE_CONFIG.FONT_SIZES.dataNumber
          : TABLE_CONFIG.FONT_SIZES.data;
        tr.setText(text);
        tr.getTextStyle()
          .setFontFamily("Montserrat").setFontSize(fontSize).setForegroundColor("#ffffff");
        tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        applyLinks(tr, getRichCell(r, c), getCellFormula(r, c));
      }
    }
  }

  for (var c = 0; c < colsCount; c++) {
    batchRequests.push({
      updateTableColumnProperties: {
        objectId: tid, columnIndices: [c],
        tableColumnProperties: { columnWidth: { magnitude: getColumnWidth(c), unit: "PT" } },
        fields: "columnWidth"
      }
    });
  }

  for (var r = 0; r < rowsCount; r++) {
    batchRequests.push({
      updateTableRowProperties: {
        objectId: tid, rowIndices: [r],
        tableRowProperties: {
          minRowHeight: {
            magnitude: r === 0 ? TABLE_CONFIG.ROW_HEIGHTS.header : TABLE_CONFIG.ROW_HEIGHTS.data,
            unit: "PT"
          }
        },
        fields: "minRowHeight"
      }
    });
  }

  batchRequests.push({
    updateTableBorderProperties: {
      objectId: tid,
      tableBorderProperties: {
        tableBorderFill: { solidFill: { color: { rgbColor: { red: 1, green: 1, blue: 1 } } } },
        weight: { magnitude: 1, unit: "PT" },
        dashStyle: "SOLID"
      },
      fields: "tableBorderFill,weight,dashStyle"
    }
  });

  for (var r = 0; r < rowsCount; r++) {
    for (var c = 0; c < colsCount; c++) {
      batchRequests.push({
        updateTableCellProperties: {
          objectId: tid,
          tableRange: { location: { rowIndex: r, columnIndex: c }, rowSpan: 1, columnSpan: 1 },
          tableCellProperties: { contentAlignment: "MIDDLE" },
          fields: "contentAlignment"
        }
      });
    }
  }

  for (var i = 0; i < groups.v.length; i++) {
    var g = groups.v[i];
    if (g.er > g.sr) {
      batchRequests.push({
        mergeTableCells: {
          objectId: tid,
          tableRange: {
            location: { rowIndex: g.sr, columnIndex: g.col },
            rowSpan: g.er - g.sr + 1, columnSpan: 1
          }
        }
      });
    }
  }

  for (var i = 0; i < groups.h.length; i++) {
    var g = groups.h[i];
    if (g.ec > g.sc) {
      batchRequests.push({
        mergeTableCells: {
          objectId: tid,
          tableRange: {
            location: { rowIndex: g.row, columnIndex: g.sc },
            rowSpan: 1, columnSpan: g.ec - g.sc + 1
          }
        }
      });
    }
  }
}


function applyFormatting(batchRequests, presentationId) {
  var BATCH_SIZE = 100;
  for (var i = 0; i < batchRequests.length; i += BATCH_SIZE) {
    try {
      Slides.Presentations.batchUpdate(
        { requests: batchRequests.slice(i, Math.min(i + BATCH_SIZE, batchRequests.length)) },
        presentationId
      );
    } catch (e) {
      Logger.log(e);
    }
  }
}
