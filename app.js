const TEMPLATE_PATH = "./assets/기간제출근부.xlsx";
const DAY_NAMES = ["월", "화", "수", "목", "금", "토", "일"];
const WEEKDAY_DAYS = new Set([0, 1, 2, 3, 4]);
const WEEKEND_DAYS = new Set([5, 6]);
const DAY_COLUMNS = [3, 4, 5, 6, 8, 10, 11, 12, 14, 16];
const COL_NAMES = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];

const state = {
  holidays: [],
};

const form = document.querySelector("#timesheet-form");
const workMonthInput = document.querySelector("#work-month");
const libraryNameInput = document.querySelector("#library-name");
const workerNameInput = document.querySelector("#worker-name");
const startTimeInput = document.querySelector("#start-time");
const endTimeInput = document.querySelector("#end-time");
const writerNameInput = document.querySelector("#writer-name");
const checkerNameInput = document.querySelector("#checker-name");
const checkerLibraryInput = document.querySelector("#checker-library");
const holidayDateInput = document.querySelector("#holiday-date");
const holidayNameInput = document.querySelector("#holiday-name");
const addHolidayButton = document.querySelector("#add-holiday");
const resetButton = document.querySelector("#reset-button");
const holidayList = document.querySelector("#holiday-list");
const documentPreview = document.querySelector("#document-preview");
const message = document.querySelector("#message");

function pad2(value) {
  return String(value).padStart(2, "0");
}

function todayMonthValue() {
  const now = new Date();
  return `${now.getFullYear()}-${pad2(now.getMonth() + 1)}`;
}

function parseMonth(value) {
  const [year, month] = value.split("-").map(Number);
  return { year, month };
}

function getLastDay(year, month) {
  return new Date(year, month, 0).getDate();
}

function getMondayBasedDay(year, month, day) {
  const jsDay = new Date(year, month - 1, day).getDay();
  return (jsDay + 6) % 7;
}

function dayToPosition(day) {
  if (day >= 1 && day <= 10) return { row: 8, col: DAY_COLUMNS[day - 1] };
  if (day >= 11 && day <= 20) return { row: 18, col: DAY_COLUMNS[day - 11] };
  if (day >= 21 && day <= 30) return { row: 28, col: DAY_COLUMNS[day - 21] };
  if (day === 31) return { row: 28, col: 17 };
  throw new Error(`지원하지 않는 날짜입니다: ${day}`);
}

function address(row, col) {
  return `${COL_NAMES[col]}${row}`;
}

function normalizeNameForFile(name) {
  return name.trim().replace(/\s+/g, "");
}

function formatNameForExcel(name) {
  const compact = normalizeNameForFile(name);
  if (compact.length >= 2 && compact.length <= 3) {
    return compact.split("").join(" ");
  }
  return name.trim();
}

function selectedWorkDays() {
  return [...document.querySelectorAll('input[name="weekday"]:checked')]
    .map((input) => Number(input.value))
    .sort((a, b) => a - b);
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle("error", isError);
}

function setDefaultDates() {
  const monthValue = todayMonthValue();
  workMonthInput.value = monthValue;
  const { year, month } = parseMonth(monthValue);
  holidayDateInput.value = `${year}-${pad2(month)}-01`;
}

function pruneHolidaysForMonth() {
  const { year, month } = parseMonth(workMonthInput.value);
  const prefix = `${year}-${pad2(month)}-`;
  state.holidays = state.holidays.filter((holiday) => holiday.date.startsWith(prefix));
}

function renderHolidays() {
  holidayList.innerHTML = "";

  if (state.holidays.length === 0) {
    const empty = document.createElement("li");
    empty.className = "holiday-empty";
    empty.textContent = "등록된 휴일이 없습니다.";
    holidayList.append(empty);
    return;
  }

  for (const holiday of state.holidays) {
    const item = document.createElement("li");
    item.className = "holiday-item";

    const text = document.createElement("span");
    text.textContent = `${holiday.date} - ${holiday.name}`;

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "remove-holiday";
    remove.textContent = "삭제";
    remove.addEventListener("click", () => {
      state.holidays = state.holidays.filter((item) => item.date !== holiday.date);
      renderHolidays();
      renderSummary();
    });

    item.append(text, remove);
    holidayList.append(item);
  }
}

function addPreviewCell(row, text, options = {}) {
  const cell = document.createElement("td");
  cell.textContent = text ?? "";
  if (options.className) cell.className = options.className;
  if (options.colSpan) cell.colSpan = options.colSpan;
  if (options.rowSpan) cell.rowSpan = options.rowSpan;
  row.append(cell);
  return cell;
}

function createPreviewRow(table, cells) {
  const row = document.createElement("tr");
  for (const cell of cells) {
    addPreviewCell(row, cell.text, cell);
  }
  table.append(row);
  return row;
}

function appendMergedSpacerRow(table) {
  const row = document.createElement("tr");
  row.className = "merged-spacer-row";
  table.append(row);
  return row;
}

function previewFooterCells(labelText, signatureText) {
  return [
    { text: "", className: "blank-cell", colSpan: 4 },
    { text: labelText, className: "footer-cell", colSpan: 5 },
    { text: signatureText, className: "footer-cell signature-cell", colSpan: 4 },
    { text: "", className: "blank-cell", colSpan: 4 },
  ];
}

function previewMonthData() {
  const { year, month } = parseMonth(workMonthInput.value);
  const lastDay = getLastDay(year, month);
  const workDays = selectedWorkDays();
  const holidays = holidayMapForMonth(year, month);
  return { year, month, lastDay, workDays, workDaySet: new Set(workDays), holidays };
}

function previewDayCell(day, data, kind) {
  if (day > data.lastDay) return { text: "", className: "time-cell" };
  const dow = getMondayBasedDay(data.year, data.month, day);
  const holidayName = data.holidays.get(day);
  const isWorkDay = data.workDaySet.has(dow);

  if (kind === "date") return { text: `${day}일`, className: "date-cell" };
  if (kind === "dow") return { text: DAY_NAMES[dow], className: "dow-cell" };
  if (!isWorkDay) {
    if (kind === "sign") return { text: "×", className: "off-cell" };
    return { text: "", className: "time-cell" };
  }
  if (holidayName) {
    if (kind === "sign") return { text: holidayName, className: "holiday-cell" };
    return { text: "", className: "time-cell" };
  }
  if (kind === "start") return { text: startTimeInput.value || "", className: "time-cell" };
  if (kind === "tilde") return { text: "~", className: "time-cell" };
  if (kind === "end") return { text: endTimeInput.value || "", className: "time-cell" };
  return { text: "", className: "time-cell" };
}

function appendDayBand(table, startDay, data) {
  const dayNumbers = Array.from({ length: startDay === 21 ? 11 : 10 }, (_, index) => startDay + index);
  const rowStart = table.rows.length;
  const leftLabelSpan = startDay === 21 ? 2 : 2;
  const tailBlankCount = startDay === 21 ? 0 : 1;

  const dateCells = [{ text: "일 자", className: "label-cell", colSpan: 2 }];
  for (const day of dayNumbers) {
    dateCells.push(previewDayCell(day, data, "date"));
  }
  if (tailBlankCount) dateCells.push({ text: "", className: "blank-cell" });
  createPreviewRow(table, dateCells);

  const dowCells = [{ text: "", className: "label-cell", colSpan: 2 }];
  for (const day of dayNumbers) {
    dowCells.push(previewDayCell(day, data, "dow"));
  }
  if (tailBlankCount) dowCells.push({ text: "", className: "blank-cell" });
  createPreviewRow(table, dowCells);

  const startCells = [{ text: "근 로\n시 간", className: "side-label", rowSpan: 3, colSpan: leftLabelSpan }];
  for (const day of dayNumbers) {
    startCells.push(previewDayCell(day, data, "start"));
  }
  if (tailBlankCount) startCells.push({ text: "", className: "blank-cell" });
  createPreviewRow(table, startCells);

  const tildeCells = [];
  for (const day of dayNumbers) {
    tildeCells.push(previewDayCell(day, data, "tilde"));
  }
  if (tailBlankCount) tildeCells.push({ text: "", className: "blank-cell" });
  createPreviewRow(table, tildeCells);

  const endCells = [];
  for (const day of dayNumbers) {
    endCells.push(previewDayCell(day, data, "end"));
  }
  if (tailBlankCount) endCells.push({ text: "", className: "blank-cell" });
  createPreviewRow(table, endCells);

  const signCells = [{ text: "서\n\n명", className: "side-label", rowSpan: 5 }, { text: "근", className: "side-label worker-char-cell worker-char-top" }];
  for (const day of dayNumbers) {
    signCells.push({ ...previewDayCell(day, data, "sign"), rowSpan: 3 });
  }
  if (tailBlankCount) signCells.push({ text: "", className: "blank-cell", rowSpan: 3 });
  const signRow = createPreviewRow(table, signCells);
  signRow.classList.add("worker-char-row");

  const workerMiddleRow = createPreviewRow(table, [{ text: "로", className: "side-label worker-char-cell worker-char-middle" }]);
  workerMiddleRow.classList.add("worker-char-row");
  const workerBottomRow = createPreviewRow(table, [{ text: "자", className: "side-label worker-char-cell worker-char-bottom" }]);
  workerBottomRow.classList.add("worker-char-row");

  const confirmCells = [{ text: "확\n인", className: "side-label", rowSpan: 2 }];
  for (const day of dayNumbers) {
    confirmCells.push({ text: day <= data.lastDay ? "" : "", className: "time-cell", rowSpan: 2 });
  }
  if (tailBlankCount) confirmCells.push({ text: "", className: "blank-cell", rowSpan: 2 });
  createPreviewRow(table, confirmCells);
  appendMergedSpacerRow(table);

  table.rows[rowStart].dataset.band = String(startDay);
}

function renderDocumentPreview() {
  if (!documentPreview || !workMonthInput.value) return;

  const data = previewMonthData();
  const table = document.createElement("table");
  table.className = "excel-preview-table";

  createPreviewRow(table, [
    { text: "", className: "blank-cell", colSpan: 3 },
    { text: "기간제근로자 출근관리부", className: "title-cell", colSpan: 12 },
    { text: "", className: "blank-cell", colSpan: 2 },
  ]);
  createPreviewRow(table, [{ text: "", className: "blank-cell", colSpan: 17 }]);
  createPreviewRow(table, [
    { text: `근무기간: ${data.year}.${pad2(data.month)}.01.~ ${pad2(data.month)}.${pad2(data.lastDay)}.`, className: "period-cell", colSpan: 8 },
    { text: "", className: "blank-cell", colSpan: 9 },
  ]);
  createPreviewRow(table, [
    { text: "근무지", className: "label-cell", colSpan: 2 },
    { text: libraryNameInput.value.trim() || "작성도서관", className: "work-cell", colSpan: 2 },
    { text: "업 무", className: "label-cell", colSpan: 2 },
    { text: `자료실 업무보조 ${getModeLabel(data.workDays)}`, className: "work-cell", colSpan: 3 },
    { text: "성 명", className: "label-cell", colSpan: 2 },
    { text: formatNameForExcel(workerNameInput.value) || "근로자", className: "name-cell", colSpan: 2 },
    { text: "", className: "blank-cell", colSpan: 4 },
  ]);
  createPreviewRow(table, [{ text: "", className: "blank-cell", colSpan: 17 }]);

  appendDayBand(table, 1, data);
  appendDayBand(table, 11, data);
  appendDayBand(table, 21, data);

  createPreviewRow(table, [{ text: "", className: "blank-cell", colSpan: 17 }]);
  createPreviewRow(
    table,
    previewFooterCells(
      `작성자: ${libraryNameInput.value.trim() || "작성도서관"}`,
      `${writerNameInput.value.trim() || "작성자"} (인)`
    )
  );
  createPreviewRow(table, [{ text: "", className: "blank-cell", colSpan: 17 }]);
  createPreviewRow(
    table,
    previewFooterCells(
      `확인자: ${checkerLibraryInput.value.trim() || "확인자 도서관"} 관장`,
      `${checkerNameInput.value.trim() || "확인자"} (인)`
    )
  );

  documentPreview.replaceChildren(table);
}

function renderSummary() {
  renderDocumentPreview();
}

function addHoliday() {
  const date = holidayDateInput.value;
  const name = holidayNameInput.value.trim() || "휴일";
  if (!date) {
    setMessage("휴일 날짜를 선택해 주세요.", true);
    return;
  }

  const { year, month } = parseMonth(workMonthInput.value);
  const prefix = `${year}-${pad2(month)}-`;
  if (!date.startsWith(prefix)) {
    setMessage("휴일은 선택한 근무년월 안의 날짜만 추가할 수 있습니다.", true);
    return;
  }

  state.holidays = state.holidays.filter((holiday) => holiday.date !== date);
  state.holidays.push({ date, name });
  state.holidays.sort((a, b) => a.date.localeCompare(b.date));
  holidayNameInput.value = "";
  setMessage("");
  renderHolidays();
  renderSummary();
}

function validateInputs() {
  const workDays = selectedWorkDays();
  if (workDays.length === 0) {
    throw new Error("근무요일을 하나 이상 선택해 주세요.");
  }
  if (startTimeInput.value >= endTimeInput.value) {
    throw new Error("근무 종료 시간은 시작 시간보다 뒤여야 합니다.");
  }
}

function getModeLabel(workDays) {
  if (workDays.every((day) => WEEKEND_DAYS.has(day))) return "(주말)";
  if (workDays.every((day) => WEEKDAY_DAYS.has(day))) return "(평일)";
  return "";
}

function setCellValue(sheet, ref, value) {
  sheet.getCell(ref).value = value;
}

function setTimeCell(sheet, ref, value) {
  const cell = sheet.getCell(ref);
  cell.value = value;
  cell.numFmt = "@";
}

function setSignCell(sheet, ref, value, holiday = false) {
  const cell = sheet.getCell(ref);
  cell.value = value;
  cell.font = {
    ...(cell.font || {}),
    bold: holiday,
    color: holiday ? { argb: "FFFF0000" } : { argb: "FF000000" },
  };
}

function clearDayBlock(sheet, row, col) {
  setCellValue(sheet, address(row, col), "");
  setCellValue(sheet, address(row + 1, col), "");
  setCellValue(sheet, address(row + 2, col), "");
  setCellValue(sheet, address(row + 3, col), "");
  setCellValue(sheet, address(row + 4, col), "");
  setSignCell(sheet, address(row + 5, col), "");
  setCellValue(sheet, address(row + 8, col), "");
}

function holidayMapForMonth(year, month) {
  const byDay = new Map();
  const prefix = `${year}-${pad2(month)}-`;
  for (const holiday of state.holidays) {
    if (holiday.date.startsWith(prefix)) {
      byDay.set(Number(holiday.date.slice(-2)), holiday.name || "휴일");
    }
  }
  return byDay;
}

function fillWorkbook(workbook) {
  const sheet = workbook.worksheets[0];
  const { year, month } = parseMonth(workMonthInput.value);
  const lastDay = getLastDay(year, month);
  const workDays = selectedWorkDays();
  const workDaySet = new Set(workDays);
  const holidays = holidayMapForMonth(year, month);
  const workerNameForFile = normalizeNameForFile(workerNameInput.value);
  const libraryName = libraryNameInput.value.trim();
  const checkerLibrary = checkerLibraryInput.value.trim();

  setCellValue(sheet, "A3", `근무기간: ${year}.${pad2(month)}.01.~ ${pad2(month)}.${pad2(lastDay)}.`);
  setCellValue(sheet, "C4", libraryName);
  setCellValue(sheet, "O4", formatNameForExcel(workerNameInput.value));
  setCellValue(sheet, "I5", getModeLabel(workDays));
  setCellValue(sheet, "H39", `작성자: ${libraryName}`);
  setCellValue(sheet, "N39", `${writerNameInput.value.trim()} (인)`);
  setCellValue(sheet, "H41", `확인자: ${checkerLibrary} 관장`);
  setCellValue(sheet, "N41", `${checkerNameInput.value.trim()} (인)`);

  for (let day = 1; day <= 31; day += 1) {
    const { row, col } = dayToPosition(day);
    if (day > lastDay) {
      clearDayBlock(sheet, row, col);
      continue;
    }

    const dow = getMondayBasedDay(year, month, day);
    const holidayName = holidays.get(day);
    const isWorkDay = workDaySet.has(dow);

    setCellValue(sheet, address(row, col), `${day}일`);
    setCellValue(sheet, address(row + 1, col), DAY_NAMES[dow]);
    setCellValue(sheet, address(row + 8, col), "");

    if (isWorkDay && !holidayName) {
      setTimeCell(sheet, address(row + 2, col), startTimeInput.value);
      setCellValue(sheet, address(row + 3, col), "~");
      setTimeCell(sheet, address(row + 4, col), endTimeInput.value);
      setSignCell(sheet, address(row + 5, col), "");
    } else if (isWorkDay && holidayName) {
      setCellValue(sheet, address(row + 2, col), "");
      setCellValue(sheet, address(row + 3, col), "");
      setCellValue(sheet, address(row + 4, col), "");
      setSignCell(sheet, address(row + 5, col), holidayName, true);
    } else {
      setCellValue(sheet, address(row + 2, col), "");
      setCellValue(sheet, address(row + 3, col), "");
      setCellValue(sheet, address(row + 4, col), "");
      setSignCell(sheet, address(row + 5, col), "×");
    }
  }

  return `출근관리부_${workerNameForFile}_${year}년_${pad2(month)}월.xlsx`;
}

function downloadBuffer(buffer, filename) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

async function loadTemplateBuffer() {
  try {
    const response = await fetch(TEMPLATE_PATH);
    if (response.ok) return response.arrayBuffer();
  } catch {
    // file:// preview cannot fetch local assets, so use the embedded fallback.
  }

  if (window.TIMESHEET_TEMPLATE_BASE64) {
    return base64ToArrayBuffer(window.TIMESHEET_TEMPLATE_BASE64);
  }

  throw new Error("엑셀 템플릿 파일을 찾을 수 없습니다.");
}

async function generateExcel() {
  validateInputs();

  if (!window.ExcelJS) {
    throw new Error("Excel 생성 라이브러리를 불러오지 못했습니다. 인터넷 연결 또는 CDN 접근을 확인해 주세요.");
  }

  setMessage("템플릿을 읽는 중입니다.");
  const templateBuffer = await loadTemplateBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  const filename = fillWorkbook(workbook);
  const outputBuffer = await workbook.xlsx.writeBuffer();
  downloadBuffer(outputBuffer, filename);
  setMessage(`${filename} 파일을 생성했습니다.`);
}

function resetForm() {
  form.reset();
  document.querySelector('input[name="weekday"][value="5"]').checked = true;
  document.querySelector('input[name="weekday"][value="6"]').checked = true;
  state.holidays = [];
  setDefaultDates();
  setMessage("");
  renderHolidays();
  renderSummary();
}

form.addEventListener("input", renderSummary);
form.addEventListener("change", renderSummary);
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    setMessage("");
    await generateExcel();
  } catch (error) {
    setMessage(error.message || "엑셀 파일 생성 중 오류가 발생했습니다.", true);
  }
});

workMonthInput.addEventListener("change", () => {
  const { year, month } = parseMonth(workMonthInput.value);
  holidayDateInput.value = `${year}-${pad2(month)}-01`;
  pruneHolidaysForMonth();
  renderHolidays();
  renderSummary();
});

addHolidayButton.addEventListener("click", addHoliday);
resetButton.addEventListener("click", resetForm);

setDefaultDates();
renderHolidays();
renderSummary();
