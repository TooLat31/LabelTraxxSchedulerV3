import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

const DAY_BLOCK_WIDTH = 6;
const IMPORT_DAY_BLOCK_STARTS = [6, 12, 18, 24, 30];
const EXPORT_DAY_BLOCK_STARTS = [1, 7, 13, 19, 25];
const DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const SCHEDULE_HEADERS = ["Customer & Job #", "EST Time", "Quantity", "Footage", "Stock", "Ship Date"];
const FINISHED_HEADERS = ["Customer & Job #", "Total CRTN", "Quantity", "Skids", "Ship Via", "Cost"];
const WORKBOOK_SHEET_PATTERN = /^\d{1,2}-\d{1,2}\s+\d{1,2}-\d{1,2}$/;
const EXPORT_SECTION_GAP_ROWS = 1;

const BASE_EXPORT_SECTIONS = [
  { press: "5.1", label: "Press 5", type: "schedule" },
  { press: "6.1", label: "Press 6", type: "schedule" },
  { press: "2.1", label: "Press 2", type: "schedule" },
  { press: "1.1", label: "Press 1", type: "schedule" },
  { press: "8", label: "Screen", type: "schedule" },
  { press: "9", label: "Grafotronic", type: "schedule" },
  { press: "Rewind", label: "Rewinding", type: "schedule" },
  { press: "Extra Duties", label: "Extra Duties", type: "schedule" },
  { press: "__finished__", label: "Finished", type: "finished" },
];

const IMPORT_PRESS_MAP = new Map([
  ["press 5", "5.1"],
  ["press 5.1", "5.1"],
  ["5", "5.1"],
  ["5.1", "5.1"],
  ["press 6", "6.1"],
  ["press 6.1", "6.1"],
  ["6", "6.1"],
  ["6.1", "6.1"],
  ["press 2", "2.1"],
  ["press 2.1", "2.1"],
  ["2", "2.1"],
  ["2.1", "2.1"],
  ["press 1", "1.1"],
  ["press 1.1", "1.1"],
  ["1", "1.1"],
  ["1.1", "1.1"],
  ["screen", "8"],
  ["press 8", "8"],
  ["8", "8"],
  ["grafotronic", "9"],
  ["press 9", "9"],
  ["9", "9"],
  ["rewind", "Rewind"],
  ["rewinding", "Rewind"],
  ["extra duties", "Extra Duties"],
  ["extra duty", "Extra Duties"],
  ["finished", "__finished__"],
]);

function safeText(value) {
  return value == null ? "" : String(value).trim();
}

function parseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = safeText(value).replace(/[$,]/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDateText(value) {
  const raw = safeText(value);
  if (!raw) return null;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function isoDate(date) {
  if (!date) return "";
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
}

function addDays(date, days) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function startOfWeek(date) {
  const copy = new Date(date);
  const day = copy.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  copy.setDate(copy.getDate() + diff);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function parseSpreadsheetDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const normalized = new Date(value);
    normalized.setHours(0, 0, 0, 0);
    return normalized;
  }
  if (typeof value === "number") {
    const parts = XLSX.SSF.parse_date_code(value);
    if (parts?.y && parts?.m && parts?.d) {
      return new Date(parts.y, parts.m - 1, parts.d);
    }
  }
  const parsed = parseDateText(value);
  if (!parsed) return null;
  parsed.setHours(0, 0, 0, 0);
  return parsed;
}

function buildSheetName(weekStart) {
  const end = addDays(weekStart, 4);
  return `${weekStart.getMonth() + 1}-${weekStart.getDate()} ${end.getMonth() + 1}-${end.getDate()}`;
}

function inferDatesFromSheetName(sheetName) {
  const match = safeText(sheetName).match(/^(\d{1,2})-(\d{1,2})\s+(\d{1,2})-(\d{1,2})$/);
  if (!match) return [];
  const [, startMonth, startDay] = match;
  const year = new Date().getFullYear();
  const start = new Date(year, Number(startMonth) - 1, Number(startDay));
  return DAY_NAMES.map((_, index) => addDays(start, index));
}

function normalizeSectionLabel(value) {
  return safeText(value).split("|")[0].trim().toLowerCase();
}

function getCell(rows, rowIndex, colIndex) {
  return rows[rowIndex]?.[colIndex] ?? null;
}

function getRowSlice(rows, rowIndex, startCol) {
  return Array.from({ length: DAY_BLOCK_WIDTH }, (_, offset) => getCell(rows, rowIndex, startCol + offset));
}

function isBlankEntry(values) {
  return values.every((value) => safeText(value) === "");
}

function extractTicketNumber(text) {
  const match = safeText(text).match(/(\d{4,})(?!.*\d)/);
  return match ? match[1] : "";
}

function extractCustomerName(text, ticketNumber) {
  const raw = safeText(text);
  if (!ticketNumber) return raw;
  const index = raw.lastIndexOf(ticketNumber);
  if (index < 0) return raw;
  return safeText(raw.slice(0, index).replace(/[-:|]+$/, ""));
}

function stockPartsFromDisplay(stockDisplay) {
  const [stockNum2 = "", stockNum1 = ""] = safeText(stockDisplay)
    .split("/")
    .map((part) => safeText(part));
  return { stockNum2, stockNum1 };
}

function makeId(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function looksLikeWeeklyScheduleSheet(sheetName) {
  return WORKBOOK_SHEET_PATTERN.test(safeText(sheetName));
}

function mapImportPressLabel(value) {
  const normalized = normalizeSectionLabel(value);
  return IMPORT_PRESS_MAP.get(normalized) || "";
}

function findSections(rows) {
  const sections = [];
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const press = mapImportPressLabel(getCell(rows, rowIndex, IMPORT_DAY_BLOCK_STARTS[0]));
    if (!press) continue;
    const previous = sections[sections.length - 1];
    if (previous) previous.endRow = rowIndex - 1;
    sections.push({
      press,
      type: press === "__finished__" ? "finished" : "schedule",
      labelRow: rowIndex,
      headerRow: rowIndex + 1,
      dataStartRow: rowIndex + 2,
      endRow: rows.length - 1,
    });
  }
  return sections;
}

function buildImportedJob(existingJob, primaryText, press, values, fallbackDate) {
  const ticketNumber = extractTicketNumber(primaryText);
  const stockDisplay = safeText(values[4]) || safeText(existingJob?.stockDisplay);
  const stockParts = stockPartsFromDisplay(stockDisplay);
  const shipByDate = parseSpreadsheetDate(values[5]) || existingJob?.shipByDate || fallbackDate || null;
  const customerName = safeText(existingJob?.customerName) || extractCustomerName(primaryText, ticketNumber);

  return {
    ...(existingJob || {}),
    id: ticketNumber,
    number: ticketNumber,
    press,
    customerName,
    generalDescr: safeText(existingJob?.generalDescr),
    custPoNum: safeText(existingJob?.custPoNum),
    priority: safeText(existingJob?.priority),
    shipByDate,
    entryDate: existingJob?.entryDate || null,
    dueOnSiteDate: existingJob?.dueOnSiteDate || shipByDate || null,
    stockNum2: stockParts.stockNum2 || safeText(existingJob?.stockNum2),
    stockNum1: stockParts.stockNum1 || safeText(existingJob?.stockNum1),
    stockDisplay,
    ticketStatus: safeText(existingJob?.ticketStatus) || "Open",
    normalizedStatus: safeText(existingJob?.normalizedStatus) || "open",
    mainTool: safeText(existingJob?.mainTool),
    toolNo2: safeText(existingJob?.toolNo2),
    ticQuantity: parseNumber(values[2]) || parseNumber(existingJob?.ticQuantity),
    estFootage: parseNumber(values[3]) || parseNumber(existingJob?.estFootage),
    estPressTime: parseNumber(values[1]) || parseNumber(existingJob?.estPressTime),
    notes: safeText(existingJob?.notes),
    holdActive: !!existingJob?.holdActive,
    holdNote: safeText(existingJob?.holdNote),
    raw: existingJob?.raw || {},
  };
}

export function importWeeklyScheduleWorkbook(arrayBuffer, existingJobs = []) {
  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
    raw: true,
  });

  const jobsById = new Map(
    (Array.isArray(existingJobs) ? existingJobs : [])
      .filter((job) => safeText(job?.id || job?.number))
      .map((job) => [safeText(job.id || job.number), job])
  );
  const importedJobs = new Map(jobsById);
  const importedAssignments = [];
  const importedDayKeys = new Set();
  const weekStarts = [];
  let importedSheetCount = 0;

  workbook.SheetNames.forEach((sheetName) => {
    if (!looksLikeWeeklyScheduleSheet(sheetName)) return;
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      blankrows: true,
      defval: null,
    });
    const sections = findSections(rows);
    if (!sections.length) return;
    importedSheetCount += 1;

    const fallbackDates = inferDatesFromSheetName(sheetName);
    const dayDescriptors = IMPORT_DAY_BLOCK_STARTS.map((startCol, index) => {
      const date = parseSpreadsheetDate(getCell(rows, 2, startCol)) || fallbackDates[index] || null;
      if (date) {
        importedDayKeys.add(isoDate(date));
        if (index === 0) weekStarts.push(startOfWeek(date));
      }
      return {
        dayKey: isoDate(date),
        date,
        startCol,
      };
    });

    sections.forEach((section) => {
      if (section.type !== "schedule") return;

      dayDescriptors.forEach((day) => {
        if (!day.date || !day.dayKey) return;
        let laneOrder = 0;

        for (let rowIndex = section.dataStartRow; rowIndex <= section.endRow; rowIndex += 1) {
          const values = getRowSlice(rows, rowIndex, day.startCol);
          if (isBlankEntry(values)) continue;

          const primaryText = safeText(values[0]);
          if (!primaryText) continue;
          if (primaryText.toLowerCase() === "customer & job #") continue;

          const ticketNumber = extractTicketNumber(primaryText);
          if (!ticketNumber) {
            importedAssignments.push({
              id: makeId("manual"),
              jobId: "",
              manualTitle: primaryText,
              dayKey: day.dayKey,
              press: section.press,
              kind: "manual",
              status: "scheduled",
              createdAt: new Date().toISOString(),
              finishedAt: null,
              finishedBy: "",
              laneOrder,
            });
            laneOrder += 1;
            continue;
          }

          const existingJob = importedJobs.get(ticketNumber);
          const importedJob = buildImportedJob(existingJob, primaryText, section.press, values, day.date);
          importedJobs.set(importedJob.id, importedJob);

          importedAssignments.push({
            id: makeId("asg"),
            jobId: importedJob.id,
            dayKey: day.dayKey,
            press: section.press,
            kind: "press",
            status: "scheduled",
            createdAt: new Date().toISOString(),
            finishedAt: null,
            finishedBy: "",
            laneOrder,
          });
          laneOrder += 1;
        }
      });
    });
  });

  return {
    jobs: Array.from(importedJobs.values()),
    assignments: importedAssignments,
    importedDayKeys: Array.from(importedDayKeys.values()),
    weekStart: weekStarts.length ? weekStarts.sort((left, right) => left - right)[0] : null,
    sheetCount: importedSheetCount,
  };
}

function scheduleEntryRow(entry) {
  if (entry.kind === "manual") {
    return [safeText(entry.title), "", "", "", "", ""];
  }
  return [
    safeText(entry.label),
    entry.estTime || "",
    entry.quantity || "",
    entry.footage || "",
    safeText(entry.stock),
    entry.shipDate || "",
  ];
}

function finishedEntryRow(entry) {
  return [
    safeText(entry.label),
    entry.totalCartons || "",
    entry.quantity || "",
    entry.skids || "",
    safeText(entry.method),
    entry.cost || "",
  ];
}

function buildExportSections(lanesByDay, finishedRowsByDay, weekColumns) {
  return BASE_EXPORT_SECTIONS.filter((section) => {
    if (section.press === "Extra Duties") {
      return weekColumns.some((day) => (lanesByDay?.[day.key]?.[section.press] || []).length);
    }
    if (section.type === "finished") return true;
    return true;
  }).map((section) => ({
    ...section,
    rowCount: Math.max(
      1,
      ...weekColumns.map((day) => {
        const rowsForDay =
          section.type === "finished"
            ? finishedRowsByDay?.[day.key] || []
            : lanesByDay?.[day.key]?.[section.press] || [];
        return rowsForDay.length;
      })
    ),
  }));
}

function applyBorder(cell, border) {
  cell.border = border;
}

function manualRowStyle(title) {
  const value = safeText(title).toLowerCase();
  if (value.includes("memorial") || value.includes("holiday")) {
    return { fill: "FFFF0000", color: "FF000000" };
  }
  return { fill: "FFFFFF00", color: "FF000000" };
}

function styleSectionLabelCell(cell) {
  cell.font = { name: "Calibri", size: 20, bold: true };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}

function styleHeaderCell(cell) {
  cell.font = { name: "Calibri", size: 8, bold: true };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}

function styleBodyCell(cell, isText = false, numFmt = "") {
  cell.font = { name: "Calibri", size: 8 };
  cell.alignment = {
    horizontal: isText ? "left" : "center",
    vertical: "middle",
    wrapText: true,
  };
  if (numFmt) cell.numFmt = numFmt;
}

export async function exportWeeklyScheduleWorkbook({ weekStart, weekColumns, lanesByDay, finishedRowsByDay }) {
  const rowPlan = buildExportSections(lanesByDay, finishedRowsByDay, weekColumns);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "OpenAI Codex";
  workbook.created = new Date();
  workbook.modified = new Date();
  const sheet = workbook.addWorksheet(buildSheetName(weekStart), {
    views: [{ state: "frozen", ySplit: 3 }],
    properties: { defaultRowHeight: 16 },
  });

  sheet.columns = [
    { width: 14 },
    { width: 9 },
    { width: 10 },
    { width: 10 },
    { width: 8 },
    { width: 8 },
    { width: 14 },
    { width: 9 },
    { width: 10 },
    { width: 10 },
    { width: 8 },
    { width: 8 },
    { width: 14 },
    { width: 9 },
    { width: 10 },
    { width: 10 },
    { width: 8 },
    { width: 8 },
    { width: 14 },
    { width: 9 },
    { width: 10 },
    { width: 10 },
    { width: 8 },
    { width: 8 },
    { width: 14 },
    { width: 9 },
    { width: 10 },
    { width: 10 },
    { width: 8 },
    { width: 8 },
  ];

  sheet.mergeCells(1, 1, 1, 30);
  const titleCell = sheet.getCell(1, 1);
  titleCell.value = "DG-Labels";
  titleCell.font = { name: "Calibri", size: 18 };
  titleCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
  applyBorder(titleCell, { bottom: { style: "thin", color: { argb: "FF000000" } } });
  sheet.getRow(1).height = 24;

  weekColumns.forEach((day, index) => {
    const startCol = EXPORT_DAY_BLOCK_STARTS[index];
    const endCol = startCol + DAY_BLOCK_WIDTH - 1;
    sheet.mergeCells(2, startCol, 2, endCol);
    sheet.mergeCells(3, startCol, 3, endCol);

    const dayCell = sheet.getCell(2, startCol);
    dayCell.value = day.label;
    dayCell.font = { name: "Calibri", size: 30 };
    dayCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    sheet.getRow(2).height = 38;

    const dateCell = sheet.getCell(3, startCol);
    dateCell.value = day.date;
    dateCell.numFmt = "m/d/yy";
    dateCell.font = { name: "Calibri", size: 22, bold: true };
    dateCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    dateCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF8EAADB" },
    };
    for (let col = startCol; col <= endCol; col += 1) {
      const cell = sheet.getCell(3, col);
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF8EAADB" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
      };
    }
  });
  sheet.getRow(3).height = 30;

  let rowCursor = 4;
  rowPlan.forEach((section) => {
    weekColumns.forEach((day, index) => {
      const startCol = EXPORT_DAY_BLOCK_STARTS[index];
      const endCol = startCol + DAY_BLOCK_WIDTH - 1;
      const rowsForDay =
        section.type === "finished"
          ? finishedRowsByDay?.[day.key] || []
          : lanesByDay?.[day.key]?.[section.press] || [];

      sheet.mergeCells(rowCursor, startCol, rowCursor, endCol);
      const labelCell = sheet.getCell(rowCursor, startCol);
      labelCell.value = section.label;
      styleSectionLabelCell(labelCell);
      sheet.getRow(rowCursor).height = 30;
      for (let col = startCol; col <= endCol; col += 1) {
        const cell = sheet.getCell(rowCursor, col);
        cell.border = {
          top: { style: "thin", color: { argb: "FF000000" } },
          bottom: { style: "thin", color: { argb: "FF000000" } },
        };
      }

      const headers = section.type === "finished" ? FINISHED_HEADERS : SCHEDULE_HEADERS;
      headers.forEach((header, headerIndex) => {
        const cell = sheet.getCell(rowCursor + 1, startCol + headerIndex);
        cell.value = header;
        styleHeaderCell(cell);
      });
      sheet.getRow(rowCursor + 1).height = 16;

      for (let itemIndex = 0; itemIndex < section.rowCount; itemIndex += 1) {
        const dataRowIndex = rowCursor + 2 + itemIndex;
        const entry = rowsForDay[itemIndex];
        sheet.getRow(dataRowIndex).height = 18;
        if (!entry) continue;

        if (section.type === "schedule" && entry.kind === "manual") {
          sheet.mergeCells(dataRowIndex, startCol, dataRowIndex, endCol);
          const cell = sheet.getCell(dataRowIndex, startCol);
          cell.value = safeText(entry.title);
          cell.font = { name: "Calibri", size: 16 };
          cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
          const style = manualRowStyle(entry.title);
          for (let col = startCol; col <= endCol; col += 1) {
            const mergedCell = sheet.getCell(dataRowIndex, col);
            mergedCell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: style.fill },
            };
          }
          continue;
        }

        const values = section.type === "finished" ? finishedEntryRow(entry) : scheduleEntryRow(entry);
        values.forEach((value, valueIndex) => {
          const cell = sheet.getCell(dataRowIndex, startCol + valueIndex);
          cell.value = value;
          const shouldLeftAlign = valueIndex === 0 || valueIndex === 4 || (section.type === "finished" && valueIndex === 4);
          const shouldCommaFormat =
            (section.type === "schedule" && (valueIndex === 2 || valueIndex === 3) && typeof value === "number") ||
            (section.type === "finished" && valueIndex === 2 && typeof value === "number");
          styleBodyCell(cell, shouldLeftAlign, shouldCommaFormat ? "#,##0" : "");
        });
      }
    });

    rowCursor += 2 + section.rowCount + EXPORT_SECTION_GAP_ROWS;
  });

  return workbook.xlsx.writeBuffer();
}
