/**
 * ============================================================
 *  ОТЧЁТ О ПРИБЫЛЯХ И УБЫТКАХ ПО ПРОЕКТАМ — Google Apps Script
 * ============================================================
 *
 *  Что делает:
 *   1. Берёт из листа «Продажи» проекты за выбранный месяц.
 *   2. Находит для каждого проекта строки в листе «Расшифровка»
 *      (поступления и списания) и сопоставляет названия.
 *   3. Делит общие расходы месяца между проектами поровну.
 *   4. Формирует на листе «Отчет» таблицу: доходы, расходы,
 *      маржинальная и валовая прибыль, маржа по каждому проекту.
 *
 *  Нужные листы (названия ищутся по ключевому слову):
 *   «Продажи»     — A: проект, B: номер месяца, далее колонка с доставкой.
 *   «Расшифровка» — колонки: Месяц, Год, Проект, Списание/Поступление,
 *                   ИТОГО в рублях, Статья, Не учитывать.
 *   «Отчет»       — B1: номер месяца, C1: год. Результат с 3-й строки.
 *
 *  Допущения (см. README):
 *   - общие расходы делятся между проектами поровну;
 *   - итоговая маржа считается от общих сумм (взвешенно).
 */


// ════════════════════════════════════════════════════════════
//  НАСТРОЙКИ
// ════════════════════════════════════════════════════════════

// Слова, по которым в заголовке листа «Продажи» ищется колонка доставки.
const DELIVERY_KEYWORDS = ["доставка"];

// Значение в колонке «Не учитывать», по которому строка считается общим расходом.
const OVERHEAD_LABEL = "общие расходы";

const COLORS = {
  headerBg      : "#1a2f4e",
  headerText    : "#ffffff",
  projectBg     : "#2c5282",
  projectText   : "#ffffff",
  rowEven       : "#f7f9fc",
  rowOdd        : "#ffffff",
  subtotalBg    : "#dbeafe",
  subtotalText  : "#1e3a5f",
  subtotalBorder: "#93c5fd",
  grandBg       : "#1a2f4e",
  grandText     : "#ffffff",
  gridColor     : "#cbd5e1",
  gridDark      : "#94a3b8",
};


// ════════════════════════════════════════════════════════════
//  ШАГ 0: РАСЧЁТ ОБЩИХ РАСХОДОВ → E1
// ════════════════════════════════════════════════════════════

function calcAndSetOverhead() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetDetail = findSheet(ss, "расшифровка");
  const sheetReport = findSheet(ss, "отчет");

  if (!sheetDetail || !sheetReport) {
    const missing = [
      !sheetDetail ? "«Расшифровка»" : null,
      !sheetReport ? "«Отчет»"       : null,
    ].filter(Boolean).join(", ");
    showError(`Не найден лист: ${missing}\nПроверьте названия листов.`);
    return null;
  }

  const period = readPeriod(sheetReport);
  if (!period) return null;

  const data   = sheetDetail.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());

  const COL_MONTH  = header.findIndex(h => h === "месяц");
  const COL_YEAR   = header.findIndex(h => h === "год");
  const COL_AMOUNT = header.findIndex(h =>
    h.includes("итого в рублях") || h === "итого"
  );
  const COL_NE = header.findIndex(h => h.includes("не учитывать"));

  const notFound = [
    { idx: COL_MONTH,  name: "«Месяц»"          },
    { idx: COL_YEAR,   name: "«Год»"             },
    { idx: COL_AMOUNT, name: "«ИТОГО в рублях»"  },
    { idx: COL_NE,     name: "«Не учитывать»"    },
  ].filter(c => c.idx === -1).map(c => c.name);

  if (notFound.length > 0) {
    Logger.log(`calcAndSetOverhead: не найдены колонки — ${notFound.join(", ")}`);
    showError(`Не найдены колонки в Расшифровке:\n${notFound.join("\n")}`);
    return null;
  }

  let total       = 0;
  let matchedRows = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (Number(row[COL_MONTH]) !== period.month) continue;
    if (Number(row[COL_YEAR])  !== period.year)  continue;
    if (String(row[COL_NE]).trim().toLowerCase() !== OVERHEAD_LABEL) continue;
    const amount = parseNum(row[COL_AMOUNT]);
    if (isNaN(amount)) continue;
    total += amount;
    matchedRows++;
  }

  Logger.log(`calcAndSetOverhead: месяц=${period.month}, год=${period.year}, строк=${matchedRows}, сумма=${total}`);

  sheetReport.getRange("E1").setValue(total);
  SpreadsheetApp.flush();

  return total;
}


// ════════════════════════════════════════════════════════════
//  ГЛАВНАЯ ФУНКЦИЯ
// ════════════════════════════════════════════════════════════

function generateProfitReport() {

  // ── ШАГ 0: общие расходы → E1 ────────────────────────────────
  const totalOverhead = calcAndSetOverhead();
  if (totalOverhead === null) return;   // ошибка уже показана пользователю

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetSales  = findSheet(ss, "Продажи");
  const sheetDetail = findSheet(ss, "Расшифровка");
  const sheetReport = findSheet(ss, "Отчет");

  if (!sheetSales || !sheetDetail || !sheetReport) {
    const names = ss.getSheets().map(s => `"${s.getName()}"`).join(", ");
    showError(`Не найден один из листов.\nВсе листы: ${names}`);
    return;
  }

  const period = readPeriod(sheetReport);
  if (!period) return;
  const targetMonth = period.month;
  const targetYear  = period.year;

  // ── Шаг 1: проекты из «Продажи» + доставка ───────────────────
  const salesData   = sheetSales.getDataRange().getValues();
  const salesHeader = salesData[0].map(h => String(h).trim().toLowerCase());

  const COL_DELIVERY = findCol(salesHeader, DELIVERY_KEYWORDS);

  const projectOrder   = [];
  const uniqueSalesSet = new Set();
  const salesOrigNames = {};
  const deliveryMap    = {};   // ключ проекта → сумма доставки за месяц

  for (let i = 1; i < salesData.length; i++) {
    const project = String(salesData[i][0]).trim();
    const month   = Number(salesData[i][1]);
    if (!project || month !== targetMonth) continue;

    const key = project.toLowerCase();
    if (!uniqueSalesSet.has(key)) {
      uniqueSalesSet.add(key);
      projectOrder.push(key);
      salesOrigNames[key] = project;
      deliveryMap[key] = 0;
    }
    // Если проект встречается в месяце несколько раз, доставки суммируются
    if (COL_DELIVERY !== -1) {
      deliveryMap[key] += parseNum(salesData[i][COL_DELIVERY]);
    }
  }

  if (projectOrder.length === 0) {
    showError(`Нет проектов за месяц ${targetMonth}.`);
    return;
  }

  const projectCount = projectOrder.length;

  // ── Шаг 2: читаем «Расшифровку» ──────────────────────────────
  const detailData = sheetDetail.getDataRange().getValues();
  const dHeader    = detailData[0].map(h => String(h).trim().toLowerCase());

  const COL_PROJ   = findCol(dHeader, ["проект"]);
  const COL_TYPE   = findCol(dHeader, ["списание/поступление"]);
  const COL_AMOUNT = findCol(dHeader, ["итого в рублях", "итого"]);
  const COL_ART    = findCol(dHeader, ["статья"]);

  const missingCols = [
    { idx: COL_PROJ,   name: "«Проект»"                },
    { idx: COL_TYPE,   name: "«Списание/Поступление»"  },
    { idx: COL_AMOUNT, name: "«ИТОГО в рублях»"        },
    { idx: COL_ART,    name: "«Статья»"                },
  ].filter(c => c.idx === -1).map(c => c.name);

  if (missingCols.length > 0) {
    showError(`Не найдены колонки в Расшифровке:\n${missingCols.join("\n")}`);
    return;
  }

  // ── Шаг 3: все уникальные проекты «Расшифровки» ──────────────
  const detailProjectNames = new Set();
  for (let i = 1; i < detailData.length; i++) {
    const p = String(detailData[i][COL_PROJ]).trim();
    if (p) detailProjectNames.add(p);
  }
  const candidateList = Array.from(detailProjectNames);

  // ── Шаг 4: сопоставление названий без дублей ─────────────────
  const matchMap           = {};
  const usedDetailProjects = new Set();

  const scored = projectOrder.map(salesKey => {
    const { match, score } = findBestMatch(salesOrigNames[salesKey], candidateList);
    return { salesKey, match, score };
  });

  scored.sort((a, b) => b.score - a.score);

  scored.forEach(({ salesKey, match }) => {
    if (match && !usedDetailProjects.has(match)) {
      matchMap[salesKey] = match;
      usedDetailProjects.add(match);
    } else {
      const { match: alt } = findBestMatch(
        salesOrigNames[salesKey],
        candidateList.filter(c => !usedDetailProjects.has(c))
      );
      matchMap[salesKey] = alt || null;
      if (alt) usedDetailProjects.add(alt);
    }
  });

  // ── Шаг 5: группируем строки «Расшифровки» по проектам ───────
  const projectData = {};
  projectOrder.forEach(key => { projectData[key] = []; });

  for (let i = 1; i < detailData.length; i++) {
    const row        = detailData[i];
    const detailProj = String(row[COL_PROJ]).trim();
    if (!detailProj) continue;

    const salesKey = projectOrder.find(k => matchMap[k] === detailProj);
    if (!salesKey) continue;

    const opType  = String(row[COL_TYPE]).trim();
    const amount  = parseNum(row[COL_AMOUNT]);
    const article = String(row[COL_ART] || "").trim();

    const income  = opType === "Поступление" ? Math.abs(amount) : 0;
    const expense = opType === "Списание"    ? -Math.abs(amount) : 0;

    projectData[salesKey].push({ article, income, expense });
  }

  // ── Шаг 6: строим строки отчёта ──────────────────────────────
  // Допущение: общие расходы делятся между проектами поровну.
  const overheadPerProject = Math.abs(totalOverhead) / projectCount;

  const outputRows = [];
  const rowMeta    = [];

  let grandSumC = 0, grandSumD = 0, grandSumE = 0;
  let grandSumF = 0, grandSumG = 0;

  projectOrder.forEach((key, projIdx) => {
    const origName      = salesOrigNames[key];
    const rows          = projectData[key];
    const isEvenProject = projIdx % 2 === 0;
    const deliveryCost  = Math.abs(deliveryMap[key] || 0);

    if (rows.length === 0 && deliveryCost === 0) {
      outputRows.push([origName, "(нет данных в расшифровке)", "", "", "", "", "", "", ""]);
      rowMeta.push({ type: "project_first", even: isEvenProject });
      outputRows.push(["Итого по проекту", "",
        0, 0, 0, overheadPerProject, -overheadPerProject, "", ""]);
      rowMeta.push({ type: "subtotal" });
      grandSumF += overheadPerProject;
      grandSumG += -overheadPerProject;
      return;
    }

    rows.forEach((r, idx) => {
      outputRows.push([
        idx === 0 ? origName : "",
        r.article,
        r.income  > 0 ? r.income  : "",
        r.expense < 0 ? r.expense : "",
        "", "", "", "", ""
      ]);
      rowMeta.push({
        type: idx === 0 ? "project_first" : "article",
        even: isEvenProject,
        idx : idx,
      });
    });

    // У проекта нет строк в расшифровке, но есть доставка: показываем название
    if (rows.length === 0 && deliveryCost > 0) {
      outputRows.push([origName, "", "", "", "", "", "", "", ""]);
      rowMeta.push({ type: "project_first", even: isEvenProject });
    }

    if (deliveryCost > 0) {
      outputRows.push(["", "Доставка", "", -deliveryCost, "", "", "", "", ""]);
      rowMeta.push({ type: "article", even: isEvenProject, idx: rows.length });
    }

    const sumIncome  = rows.reduce((s, r) => s + r.income,  0);
    const sumExpense = rows.reduce((s, r) => s + r.expense, 0) - deliveryCost;

    const marginal = sumIncome + sumExpense;       // маржинальная прибыль
    const gross    = marginal - overheadPerProject; // после общих расходов
    const margin   = sumIncome !== 0 ? marginal / sumIncome : "";

    grandSumC += sumIncome;
    grandSumD += sumExpense;
    grandSumE += marginal;
    grandSumF += overheadPerProject;
    grandSumG += gross;

    outputRows.push([
      "Итого по проекту", "",
      sumIncome, sumExpense,
      marginal, overheadPerProject, gross, margin, ""
    ]);
    rowMeta.push({ type: "subtotal" });
  });

  // Итоговая маржа считается от общих сумм, а не как среднее по проектам
  const grandMargin = grandSumC !== 0 ? grandSumE / grandSumC : "";

  outputRows.push([
    "ИТОГО", "",
    grandSumC, grandSumD,
    grandSumE, grandSumF,
    grandSumG, grandMargin, ""
  ]);
  rowMeta.push({ type: "grand" });

  // ── Шаг 7: очищаем лист и записываем ─────────────────────────
  const lastRow = sheetReport.getLastRow();
  if (lastRow >= 3) {
    sheetReport.getRange(3, 1, lastRow - 2, 9)
      .clearContent().setBackground(null).setFontWeight("normal")
      .setFontColor(null).setFontSize(10)
      .setBorder(false, false, false, false, false, false)
      .setHorizontalAlignment("left");
    SpreadsheetApp.flush();
  }

  sheetReport.getRange("I1").setValue(projectCount);
  sheetReport.getRange("I3").setValue(projectCount);
  sheetReport.getRange(3, 1, outputRows.length, 9).setValues(outputRows);
  SpreadsheetApp.flush();

  // ── Шаг 8: форматирование ────────────────────────────────────
  formatReport(sheetReport, outputRows.length, rowMeta);

  // Лог сопоставлений (Расширения → Apps Script → Выполнение)
  Logger.log("=== Сопоставления проектов ===");
  projectOrder.forEach(k =>
    Logger.log(`  "${salesOrigNames[k]}" → "${matchMap[k] || "НЕ НАЙДЕН"}" | доставка: ${deliveryMap[k] || 0}`)
  );

  const unmatched = projectOrder.filter(k => !matchMap[k]).map(k => salesOrigNames[k]);
  const months = ["", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];

  SpreadsheetApp.getUi().alert(
    `✅ Отчёт сформирован\n\n` +
    `Период: ${months[targetMonth]} ${targetYear}\n` +
    `Общие расходы (E1): ${formatMoney(totalOverhead)} руб.\n` +
    `Проектов: ${projectCount}\n` +
    `Общие расходы на проект: ${formatMoney(overheadPerProject)} руб.\n` +
    `Строк в отчёте: ${outputRows.length}` +
    (unmatched.length
      ? `\n\n⚠️ Не найдены в расшифровке:\n${unmatched.join("\n")}`
      : "")
  );
}


// ════════════════════════════════════════════════════════════
//  ФОРМАТИРОВАНИЕ
// ════════════════════════════════════════════════════════════

function formatReport(sheet, totalRows, rowMeta) {
  const startRow = 3;
  const numFmt   = '#,##0.00;[RED]-#,##0.00;"-"';
  const pctFmt   = "0.00%";

  sheet.getRange(startRow, 1, totalRows, 9)
    .setFontFamily("Arial").setFontSize(10).setVerticalAlignment("middle");

  rowMeta.forEach((meta, idx) => {
    const r   = startRow + idx;
    const row = sheet.getRange(r, 1, 1, 9);

    switch (meta.type) {
      case "project_first":
        row.setBackground(COLORS.projectBg)
           .setFontColor(COLORS.projectText)
           .setFontWeight("bold");
        row.setBorder(true, null, null, null, null, null,
          COLORS.headerBg, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(r, 3, 1, 2).setNumberFormat(numFmt);
        break;

      case "article":
        row.setBackground(meta.idx % 2 === 0 ? COLORS.rowOdd : COLORS.rowEven)
           .setFontColor("#374151");
        row.setBorder(false, null, true, null, null, null,
          COLORS.gridColor, SpreadsheetApp.BorderStyle.DOTTED);
        sheet.getRange(r, 3, 1, 2).setNumberFormat(numFmt);
        break;

      case "subtotal":
        row.setBackground(COLORS.subtotalBg)
           .setFontColor(COLORS.subtotalText)
           .setFontWeight("bold");
        row.setBorder(true, null, true, null, null, null,
          COLORS.subtotalBorder, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(r, 3, 1, 4).setNumberFormat(numFmt);
        sheet.getRange(r, 8, 1, 1).setNumberFormat(pctFmt);
        break;

      case "grand":
        row.setBackground(COLORS.grandBg)
           .setFontColor(COLORS.grandText)
           .setFontWeight("bold").setFontSize(11);
        row.setBorder(true, null, true, null, null, null,
          "#0f172a", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.getRange(r, 3, 1, 5).setNumberFormat(numFmt);
        sheet.getRange(r, 8, 1, 1).setNumberFormat(pctFmt);
        break;
    }
  });

  sheet.getRange(startRow, 3, totalRows, 6).setHorizontalAlignment("right");
  sheet.getRange(startRow, 1, totalRows, 2).setHorizontalAlignment("left");

  // Внешняя рамка и вертикальные линии
  sheet.getRange(startRow, 1, totalRows, 9)
    .setBorder(true, true, true, true, null, null,
      COLORS.gridDark, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(startRow, 1, totalRows, 9)
    .setBorder(null, null, null, null, true, null,
      COLORS.gridColor, SpreadsheetApp.BorderStyle.SOLID);

  // Заголовки
  sheet.getRange(2, 1, 1, 9)
    .setBackground(COLORS.headerBg)
    .setFontColor(COLORS.headerText)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setRowHeight(2, 40);

  rowMeta.forEach((meta, idx) => {
    sheet.setRowHeight(startRow + idx, meta.type === "article" ? 22 : 28);
  });

  [200, 220, 130, 130, 140, 140, 140, 110, 100].forEach((w, i) =>
    sheet.setColumnWidth(i + 1, w)
  );

  sheet.setFrozenRows(2);
}


// ════════════════════════════════════════════════════════════
//  СОПОСТАВЛЕНИЕ НАЗВАНИЙ ПРОЕКТОВ
// ════════════════════════════════════════════════════════════

function findBestMatch(salesName, candidateList) {
  if (!candidateList || candidateList.length === 0) {
    return { match: null, score: 0 };
  }
  // 1. точное совпадение
  for (const c of candidateList) {
    if (c.toLowerCase().trim() === salesName.toLowerCase().trim()) {
      return { match: c, score: 999 };
    }
  }
  // 2. совпадение по четырёхзначным номерам + пересечение слов
  const digits4 = extractDigits4(salesName);
  let bestMatch = null;
  let bestScore = -1;

  for (const det of candidateList) {
    const dDigits      = extractDigits4(det);
    const digitMatches = digits4.filter(d => dDigits.includes(d)).length;
    if (digitMatches === 0 && digits4.length > 0) continue;
    const score = digitMatches * 100 + wordOverlapScore(salesName, det);
    if (score > bestScore) { bestScore = score; bestMatch = det; }
  }

  // 3. в названии нет номеров: сравниваем только слова
  if (!bestMatch && digits4.length === 0) {
    for (const det of candidateList) {
      const ws = wordOverlapScore(salesName, det);
      if (ws > bestScore) { bestScore = ws; bestMatch = det; }
    }
  }
  return { match: bestMatch, score: bestScore };
}

function extractDigits4(name) {
  const m = String(name).match(/\b\d{4}\b/g);
  return m ? m : [];
}

function wordOverlapScore(a, b) {
  const wa = tokenize(a);
  const wb = tokenize(b);
  if (!wa.length || !wb.length) return 0;
  const sb = new Set(wb);
  return Math.round(
    wa.filter(w => sb.has(w)).length / Math.max(wa.length, wb.length) * 100
  );
}

function tokenize(str) {
  return String(str).toLowerCase()
    .replace(/\d+/g, " ")
    .split(/[\s\-\/]+/)
    .map(w => w.trim())
    .filter(w => w.length > 1);
}


// ════════════════════════════════════════════════════════════
//  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ════════════════════════════════════════════════════════════

/** Месяц и год из B1 и C1 листа «Отчет»; при ошибке показывает сообщение. */
function readPeriod(sheetReport) {
  const month = Number(sheetReport.getRange("B1").getValue());
  const year  = Number(sheetReport.getRange("C1").getValue());

  if (!month || isNaN(month) || month < 1 || month > 12) {
    showError("Ячейка B1 пуста или не число.\nВведите номер месяца от 1 до 12.");
    return null;
  }
  if (!year || isNaN(year) || year < 2000) {
    showError("Ячейка C1 пуста или не число.\nВведите год, например 2026.");
    return null;
  }
  return { month, year };
}

function findSheet(ss, keyword) {
  const kw = keyword.toLowerCase();
  return ss.getSheets().find(s => s.getName().toLowerCase().includes(kw)) || null;
}

function findCol(headers, keywords) {
  for (let i = 0; i < headers.length; i++) {
    for (const kw of keywords) {
      if (headers[i].includes(kw.toLowerCase())) return i;
    }
  }
  return -1;
}

function parseNum(val) {
  if (typeof val === "number") return val;
  const n = parseFloat(String(val).replace(/\s/g, "").replace(",", "."));
  return isNaN(n) ? 0 : n;
}

function formatMoney(n) {
  return n.toLocaleString("ru-RU", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function showError(msg) {
  SpreadsheetApp.getUi().alert("⚠️ Ошибка\n\n" + msg);
}


// ════════════════════════════════════════════════════════════
//  МЕНЮ
// ════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Отчёты")
    .addItem("▶ Сформировать отчёт", "generateProfitReport")
    .addToUi();
}
