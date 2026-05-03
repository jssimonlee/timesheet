const TEMPLATE_PATH = "./assets/기간제출근부.xlsx";
const DAY_NAMES = ["월", "화", "수", "목", "금", "토", "일"];
const WEEKDAY_DAYS = new Set([0, 1, 2, 3, 4]);
const WEEKEND_DAYS = new Set([5, 6]);
const DAY_COLUMNS = [3, 4, 5, 6, 8, 10, 11, 12, 14, 16];
const COL_NAMES = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];
const PREVIEW_DAY_SLOT_WIDTH = 43.17 * 0.8;
const PREVIEW_HALF_DAY_SLOT_WIDTH = PREVIEW_DAY_SLOT_WIDTH / 2;
const PREVIEW_COLUMN_WIDTHS = [
  19,
  19,
  PREVIEW_DAY_SLOT_WIDTH,
  PREVIEW_DAY_SLOT_WIDTH,
  PREVIEW_DAY_SLOT_WIDTH,
  PREVIEW_HALF_DAY_SLOT_WIDTH,
  PREVIEW_HALF_DAY_SLOT_WIDTH,
  PREVIEW_HALF_DAY_SLOT_WIDTH,
  PREVIEW_HALF_DAY_SLOT_WIDTH,
  PREVIEW_DAY_SLOT_WIDTH,
  PREVIEW_DAY_SLOT_WIDTH,
  PREVIEW_HALF_DAY_SLOT_WIDTH,
  PREVIEW_HALF_DAY_SLOT_WIDTH,
  PREVIEW_HALF_DAY_SLOT_WIDTH,
  PREVIEW_HALF_DAY_SLOT_WIDTH,
  PREVIEW_DAY_SLOT_WIDTH,
  PREVIEW_DAY_SLOT_WIDTH,
];
const STORAGE_KEY = "timesheet-form-defaults";
const HOLIDAY_STORAGE_KEY = "timesheet-holiday-settings";
const DEFAULT_WORK_DAYS = [5, 6];
const DEFAULT_START_TIME = "10:00";
const DEFAULT_END_TIME = "18:00";
const WEEKDAY_PRESET = {
  days: [0, 1, 2, 3, 4],
  startTime: "10:00",
  endTime: "13:00",
};
const WEEKEND_EARLY_PRESET = {
  days: [5, 6],
  startTime: "09:00",
  endTime: "17:00",
};
const WEEKEND_PRESET = {
  days: [5, 6],
  startTime: "10:00",
  endTime: "18:00",
};

const state = {
  customHolidays: [],
  deletedDefaultHolidays: [],
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
const holidayModal = document.querySelector("#holiday-modal");
const openHolidayModalButton = document.querySelector("#open-holiday-modal-btn");
const closeHolidayModalButton = document.querySelector("#close-holiday-modal-btn");
const weekdayPresetButton = document.querySelector("#weekday-preset-button");
const weekendEarlyPresetButton = document.querySelector("#weekend-early-preset-button");
const weekendPresetButton = document.querySelector("#weekend-preset-button");
const holidayDateInput = document.querySelector("#holiday-date");
const holidayNameInput = document.querySelector("#holiday-name");
const addHolidayButton = document.querySelector("#add-holiday");
const holidayList = document.querySelector("#holiday-list");
const documentPreview = document.querySelector("#document-preview");
const message = document.querySelector("#message");
const WORK_PRESETS = [
  { button: weekdayPresetButton, preset: WEEKDAY_PRESET },
  { button: weekendEarlyPresetButton, preset: WEEKEND_EARLY_PRESET },
  { button: weekendPresetButton, preset: WEEKEND_PRESET },
];

function pad2(value) {
  return String(value).padStart(2, "0");
}

function todayMonthValue() {
  const nextMonth = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1);
  return `${nextMonth.getFullYear()}-${pad2(nextMonth.getMonth() + 1)}`;
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

function formatShortNameForExcel(name) {
  const compact = normalizeNameForFile(name);
  if (compact.length >= 1 && compact.length <= 3) {
    return compact.split("").join(" ");
  }
  return name.trim();
}

function selectedWorkDays() {
  return [...document.querySelectorAll('input[name="weekday"]:checked')]
    .map((input) => Number(input.value))
    .sort((a, b) => a - b);
}

function setSelectedWorkDays(days) {
  const selected = new Set((Array.isArray(days) && days.length ? days : DEFAULT_WORK_DAYS).map(Number));

  for (const input of document.querySelectorAll('input[name="weekday"]')) {
    input.checked = selected.has(Number(input.value));
  }
}

function sameWorkDays(left, right) {
  if (left.length !== right.length) return false;
  return left.every((day, index) => day === right[index]);
}

function matchesWorkPreset(preset) {
  return (
    sameWorkDays(selectedWorkDays(), preset.days) &&
    startTimeInput.value === preset.startTime &&
    endTimeInput.value === preset.endTime
  );
}

function syncWorkPresetButtons() {
  for (const { button, preset } of WORK_PRESETS) {
    const isActive = matchesWorkPreset(preset);
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  }
}

function applyWorkPreset(preset) {
  setSelectedWorkDays(preset.days);
  startTimeInput.value = preset.startTime;
  endTimeInput.value = preset.endTime;
  writeStoredFormDefaults();
  renderSummary();
}

function readStoredFormDefaults() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};

    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function writeStoredFormDefaults() {
  try {
    window.localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        libraryName: libraryNameInput.value.trim(),
        startTime: startTimeInput.value || DEFAULT_START_TIME,
        endTime: endTimeInput.value || DEFAULT_END_TIME,
        writerName: writerNameInput.value.trim(),
        checkerName: checkerNameInput.value.trim(),
        checkerLibrary: checkerLibraryInput.value.trim(),
        workDays: selectedWorkDays(),
      })
    );
  } catch {
    // Ignore storage failures in restricted environments.
  }
}

function applyStoredFormDefaults() {
  const stored = readStoredFormDefaults();
  libraryNameInput.value = stored.libraryName ?? "";
  startTimeInput.value = stored.startTime || DEFAULT_START_TIME;
  endTimeInput.value = stored.endTime || DEFAULT_END_TIME;
  writerNameInput.value = stored.writerName ?? "";
  checkerNameInput.value = stored.checkerName ?? "";
  checkerLibraryInput.value = stored.checkerLibrary ?? "";
  setSelectedWorkDays(stored.workDays);
}

function readStoredHolidaySettings() {
  try {
    const raw = window.localStorage.getItem(HOLIDAY_STORAGE_KEY);
    if (!raw) return {};

    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function writeStoredHolidaySettings() {
  try {
    window.localStorage.setItem(
      HOLIDAY_STORAGE_KEY,
      JSON.stringify({
        customHolidays: state.customHolidays,
        deletedDefaultHolidays: state.deletedDefaultHolidays,
      })
    );
  } catch {
    // Ignore storage failures in restricted environments.
  }
}

function applyStoredHolidaySettings() {
  const stored = readStoredHolidaySettings();
  const customByDate = new Map();

  for (const holiday of Array.isArray(stored.customHolidays) ? stored.customHolidays : []) {
    if (!holiday || typeof holiday.date !== "string") continue;
    customByDate.set(holiday.date, {
      date: holiday.date,
      name: typeof holiday.name === "string" && holiday.name.trim() ? holiday.name.trim() : "휴일",
    });
  }

  state.customHolidays = [...customByDate.values()].sort((a, b) => a.date.localeCompare(b.date));
  state.deletedDefaultHolidays = [
    ...new Set(
      (Array.isArray(stored.deletedDefaultHolidays) ? stored.deletedDefaultHolidays : []).filter(
        (date) => typeof date === "string"
      )
    ),
  ];
}

function setMessage(text, isError = false) {
  message.textContent = text;
  message.classList.toggle("error", isError);
}

function openHolidayModal() {
  syncHolidaysForMonth();
  renderHolidays();
  holidayModal.hidden = false;
  requestAnimationFrame(() => {
    holidayDateInput.focus();
  });
}

function closeHolidayModal() {
  holidayModal.hidden = true;
}

function setDefaultDates() {
  const monthValue = todayMonthValue();
  workMonthInput.value = monthValue;
}

function defaultHolidaysForMonth(year, month) {
  const prefix = `${year}-${pad2(month)}-`;
  const deletedDefaults = new Set(state.deletedDefaultHolidays);

  return (Array.isArray(window.DEFAULT_HOLIDAYS) ? window.DEFAULT_HOLIDAYS : [])
    .filter((holiday) => holiday.date.startsWith(prefix) && !deletedDefaults.has(holiday.date))
    .map((holiday) => ({
      date: holiday.date,
      name: holiday.name || "휴일",
      source: "default",
    }));
}

function combinedHolidaysForMonth(year, month) {
  const prefix = `${year}-${pad2(month)}-`;
  const holidaysByDate = new Map(defaultHolidaysForMonth(year, month).map((holiday) => [holiday.date, holiday]));

  for (const holiday of state.customHolidays) {
    if (holiday.date.startsWith(prefix)) {
      holidaysByDate.set(holiday.date, {
        ...holiday,
        source: "custom",
      });
    }
  }

  return [...holidaysByDate.values()].sort((a, b) => a.date.localeCompare(b.date));
}

function syncHolidaysForMonth() {
  if (!workMonthInput.value) return;

  const { year, month } = parseMonth(workMonthInput.value);
  holidayDateInput.value = `${year}-${pad2(month)}-01`;
  state.holidays = combinedHolidaysForMonth(year, month);
}

function removeHoliday(holiday) {
  if (holiday.source === "default") {
    if (!state.deletedDefaultHolidays.includes(holiday.date)) {
      state.deletedDefaultHolidays.push(holiday.date);
    }
  } else {
    state.customHolidays = state.customHolidays.filter((item) => item.date !== holiday.date);
  }

  writeStoredHolidaySettings();
  syncHolidaysForMonth();
  renderHolidays();
  renderSummary();
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

    const meta = document.createElement("div");
    meta.className = "holiday-item-meta";

    const date = document.createElement("span");
    date.className = "holiday-date-label";
    date.textContent = holiday.date;

    const nameRow = document.createElement("div");
    nameRow.className = "holiday-name-row";

    const name = document.createElement("span");
    name.className = "holiday-name-label";
    name.textContent = holiday.name;

    const sourceTag = document.createElement("span");
    sourceTag.className = `holiday-source-tag ${holiday.source === "default" ? "is-default" : "is-custom"}`;
    sourceTag.textContent = holiday.source === "default" ? "기본" : "추가됨";

    nameRow.append(name, sourceTag);
    meta.append(date, nameRow);

    item.append(meta);

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "remove-holiday";
    remove.innerHTML = "&times;";
    remove.title = `${holiday.name} 삭제`;
    remove.setAttribute("aria-label", `${holiday.name} 삭제`);
    remove.addEventListener("click", () => {
      removeHoliday(holiday);
    });

    item.append(remove);

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

function appendPreviewColumnGroup(table) {
  const colGroup = document.createElement("colgroup");
  let totalWidth = 0;

  for (const width of PREVIEW_COLUMN_WIDTHS) {
    const col = document.createElement("col");
    col.style.width = `${width}px`;
    totalWidth += width;
    colGroup.append(col);
  }

  table.append(colGroup);
  table.style.width = `${totalWidth}px`;
  table.style.margin = "0 auto";
}

function appendMergedSpacerRow(table) {
  const row = document.createElement("tr");
  row.className = "merged-spacer-row";
  table.append(row);
  return row;
}

function previewFooterCells(labelText, signatureText) {
  return [
    { text: "", className: "blank-cell", colSpan: 7 },
    { text: labelText, className: "footer-cell", colSpan: 6 },
    { text: signatureText, className: "footer-cell signature-cell", colSpan: 4 },
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

function previewBandSlots(startDay) {
  const dayCount = startDay === 21 ? 11 : 10;
  const dayNumbers = Array.from({ length: dayCount }, (_, index) => startDay + index);
  const columnPositions = startDay === 21 ? [...DAY_COLUMNS, 17] : DAY_COLUMNS;
  const bandEndColumn = startDay === 21 ? 18 : 17;
  const slots = dayNumbers.map((day, index) => ({
    day,
    colSpan: (columnPositions[index + 1] ?? bandEndColumn) - columnPositions[index],
  }));

  if (startDay !== 21) {
    slots.push({ day: null, colSpan: 1, className: "tail-fill-cell" });
  }

  return slots;
}

function appendDayBand(table, startDay, data) {
  const daySlots = previewBandSlots(startDay);
  const rowStart = table.rows.length;
  const leftLabelSpan = startDay === 21 ? 2 : 2;

  const dateCells = [{ text: "일 자", className: "label-cell", colSpan: 2 }];
  for (const slot of daySlots) {
    dateCells.push(
      slot.day === null
        ? { text: "", className: slot.className, colSpan: slot.colSpan }
        : { ...previewDayCell(slot.day, data, "date"), colSpan: slot.colSpan }
    );
  }
  createPreviewRow(table, dateCells);

  const dowCells = [{ text: "", className: "label-cell", colSpan: 2 }];
  for (const slot of daySlots) {
    dowCells.push(
      slot.day === null
        ? { text: "", className: slot.className, colSpan: slot.colSpan }
        : { ...previewDayCell(slot.day, data, "dow"), colSpan: slot.colSpan }
    );
  }
  createPreviewRow(table, dowCells);

  const startCells = [{ text: "근 로\n시 간", className: "side-label", rowSpan: 3, colSpan: leftLabelSpan }];
  for (const slot of daySlots) {
    startCells.push(
      slot.day === null
        ? { text: "", className: slot.className, colSpan: slot.colSpan, rowSpan: 3 }
        : { ...previewDayCell(slot.day, data, "start"), colSpan: slot.colSpan }
    );
  }
  const startRow = createPreviewRow(table, startCells);
  startRow.classList.add("time-start-row");

  const tildeCells = [];
  for (const slot of daySlots) {
    if (slot.day === null) continue;
    tildeCells.push({ ...previewDayCell(slot.day, data, "tilde"), colSpan: slot.colSpan });
  }
  const middleRow = createPreviewRow(table, tildeCells);
  middleRow.classList.add("time-middle-row");

  const endCells = [];
  for (const slot of daySlots) {
    if (slot.day === null) continue;
    endCells.push({ ...previewDayCell(slot.day, data, "end"), colSpan: slot.colSpan });
  }
  const endRow = createPreviewRow(table, endCells);
  endRow.classList.add("time-end-row");

  const signCells = [{ text: "서\n\n명", className: "side-label", rowSpan: 5 }, { text: "근", className: "side-label worker-char-cell worker-char-top" }];
  for (const slot of daySlots) {
    signCells.push(
      slot.day === null
        ? { text: "", className: slot.className, colSpan: slot.colSpan, rowSpan: 3 }
        : { ...previewDayCell(slot.day, data, "sign"), colSpan: slot.colSpan, rowSpan: 3 }
    );
  }
  const signRow = createPreviewRow(table, signCells);
  signRow.classList.add("worker-char-row");

  const workerMiddleRow = createPreviewRow(table, [{ text: "로", className: "side-label worker-char-cell worker-char-middle" }]);
  workerMiddleRow.classList.add("worker-char-row");
  const workerBottomRow = createPreviewRow(table, [{ text: "자", className: "side-label worker-char-cell worker-char-bottom" }]);
  workerBottomRow.classList.add("worker-char-row");

  const confirmCells = [{ text: "확\n인", className: "side-label", rowSpan: 2 }];
  for (const slot of daySlots) {
    confirmCells.push(
      slot.day === null
        ? { text: "", className: slot.className, colSpan: slot.colSpan, rowSpan: 2 }
        : { text: "", className: "time-cell", colSpan: slot.colSpan, rowSpan: 2 }
    );
  }
  const confirmRow = createPreviewRow(table, confirmCells);
  confirmRow.classList.add("confirm-row");
  appendMergedSpacerRow(table);

  table.rows[rowStart].dataset.band = String(startDay);
}

function renderDocumentPreview() {
  if (!documentPreview || !workMonthInput.value) return;

  const data = previewMonthData();
  const table = document.createElement("table");
  table.className = "excel-preview-table";
  appendPreviewColumnGroup(table);

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
    { text: libraryNameInput.value.trim() || "작성도서관", className: "work-cell", colSpan: 3 },
    { text: "업 무", className: "label-cell", colSpan: 2 },
    { text: `자료실 업무보조 ${getModeLabel(data.workDays)}`, className: "work-cell", colSpan: 4 },
    { text: "성 명", className: "label-cell", colSpan: 2 },
    { text: formatNameForExcel(workerNameInput.value) || "근로자", className: "name-cell", colSpan: 4 },
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
      `${formatShortNameForExcel(writerNameInput.value) || "작성자"} (인)`
    )
  );
  createPreviewRow(table, [{ text: "", className: "footer-gap-cell", colSpan: 17 }]);
  createPreviewRow(
    table,
    previewFooterCells(
      `확인자: ${checkerLibraryInput.value.trim() || "확인자 도서관"} 관장`,
      `${formatShortNameForExcel(checkerNameInput.value) || "확인자"} (인)`
    )
  );

  documentPreview.replaceChildren(table);
}

function renderSummary() {
  syncWorkPresetButtons();
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

  state.deletedDefaultHolidays = state.deletedDefaultHolidays.filter((item) => item !== date);
  state.customHolidays = state.customHolidays.filter((holiday) => holiday.date !== date);
  state.customHolidays.push({ date, name });
  state.customHolidays.sort((a, b) => a.date.localeCompare(b.date));
  writeStoredHolidaySettings();
  holidayNameInput.value = "";
  setMessage("");
  syncHolidaysForMonth();
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

function cloneStyleValue(value) {
  if (!value || typeof value !== "object") return value;
  if (Array.isArray(value)) return value.map(cloneStyleValue);

  return Object.fromEntries(
    Object.entries(value).map(([key, entry]) => [key, cloneStyleValue(entry)])
  );
}

function detachCellStyle(cell) {
  cell.style = cloneStyleValue(cell.style || {});
}

function setTimeCell(sheet, ref, value) {
  const cell = sheet.getCell(ref);
  detachCellStyle(cell);
  cell.value = value;
  cell.numFmt = "@";
}

function setSignCell(sheet, ref, value, holiday = false) {
  const cell = sheet.getCell(ref);
  detachCellStyle(cell);
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
  for (const holiday of combinedHolidaysForMonth(year, month)) {
    byDay.set(Number(holiday.date.slice(-2)), holiday.name || "휴일");
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
  setCellValue(sheet, "N39", `${formatShortNameForExcel(writerNameInput.value) || "작성자"} (인)`);
  setCellValue(sheet, "H41", `확인자: ${checkerLibrary} 관장`);
  setCellValue(sheet, "N41", `${formatShortNameForExcel(checkerNameInput.value) || "확인자"} (인)`);

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

form.addEventListener("input", () => {
  writeStoredFormDefaults();
  renderSummary();
});
form.addEventListener("change", () => {
  writeStoredFormDefaults();
  renderSummary();
});
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
  syncHolidaysForMonth();
  renderHolidays();
  renderSummary();
});

openHolidayModalButton.addEventListener("click", openHolidayModal);
closeHolidayModalButton.addEventListener("click", closeHolidayModal);
holidayModal.addEventListener("click", (event) => {
  if (event.target === holidayModal) {
    closeHolidayModal();
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !holidayModal.hidden) {
    closeHolidayModal();
  }
});

weekdayPresetButton.addEventListener("click", () => {
  applyWorkPreset(WEEKDAY_PRESET);
});

weekendEarlyPresetButton.addEventListener("click", () => {
  applyWorkPreset(WEEKEND_EARLY_PRESET);
});

weekendPresetButton.addEventListener("click", () => {
  applyWorkPreset(WEEKEND_PRESET);
});

addHolidayButton.addEventListener("click", addHoliday);

setDefaultDates();
applyStoredFormDefaults();
applyStoredHolidaySettings();
workerNameInput.value = "";
syncHolidaysForMonth();
renderHolidays();
renderSummary();
