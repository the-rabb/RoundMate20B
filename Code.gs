const DEFAULT_DEPARTMENTS = [
  { name: "แผนก 1", beds: ["1", "2", "3", "4", "5", "6", "7", "8"] },
  { name: "แผนก 2", beds: ["1", "2", "3", "4", "5", "6", "7"] },
  { name: "แผนก 3", beds: ["1", "2", "3", "4", "5", "6"] },
  { name: "แผนก 4", beds: ["1", "2", "3", "4", "5", "6"] },
  { name: "ห้องแยก", beds: ["1", "2"] },
];

const DEFAULT_EMPTY_REASONS = [
  "ไม่มีผู้ป่วย", "ผู้ป่วยรอกลับบ้าน", "รอรับผู้ป่วยกลับจาก OR", "ผู้ป่วยไปทำ CT/MRI/Xray", "ผู้ป่วยรอย้ายตึก"
];

const DEFAULT_DRAIN_TYPES = [
  "ETT/TT", "Traction", "NG/OG", "F/C", "JD", "RD", "PCD/PTBD", "ICD", "VAC dressing", "EVD", "Subdural Drain", "Lumbar drain", "N/A"
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("แบบบันทึกการตรวจเยี่ยมผู้ป่วยตามหลัก I-SITDOWN")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName("Setting");
  if (!settingSheet) throw new Error("ไม่พบแท็บ Setting");

  const rows = settingSheet.getDataRange().getValues();
  if (rows.length === 0) throw new Error("แท็บ Setting ไม่มีข้อมูล");

  const headers = rows.shift().map(String);
  const nameIdx = headers.indexOf("StaffName");
  const activeIdx = headers.indexOf("IsActice");
  const drainIdx = headers.indexOf("DrainTypes");

  if (nameIdx < 0 || activeIdx < 0) {
    throw new Error("แท็บ Setting ต้องมีคอลัมน์ StaffName และ IsActice");
  }

  const staff = [];
  const drainTypes = [];

  rows.forEach(row => {
    const sName = row[nameIdx];
    const sActive = row[activeIdx];
    if (sName && (sActive === true || String(sActive).toUpperCase() === "TRUE" || String(sActive) === "1")) {
      staff.push(String(sName).trim());
    }

    if (drainIdx >= 0 && row[drainIdx]) {
      const dVal = String(row[drainIdx]).trim();
      if (dVal && !drainTypes.includes(dVal)) {
        drainTypes.push(dVal);
      }
    }
  });

  return {
    staff: staff,
    departments: DEFAULT_DEPARTMENTS,
    emptyReasons: DEFAULT_EMPTY_REASONS,
    drainTypes: drainTypes.length > 0 ? drainTypes : DEFAULT_DRAIN_TYPES,
    shift: getShiftInfo_(),
  };
}

function saveRound(payload) {
  if (!payload || !payload.date || !payload.shift || !payload.nurseName) {
    throw new Error("ข้อมูลวันที่ เวร หรือผู้ประเมินไม่ครบ");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("dataISITDOWN");

  if (!sheet) {
    sheet = ss.insertSheet("dataISITDOWN");
    sheet.appendRow(["data", "shift", "nurseName", "completeAt"]);
  }

  const dateParts = payload.date.split("-").map(Number);
  if (dateParts.length !== 3 || dateParts.some(n => !n)) {
    throw new Error("รูปแบบวันที่ไม่ถูกต้อง");
  }

  const roundDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2], 12, 0, 0);
  const completedAt = new Date();

  sheet.appendRow([roundDate, payload.shift, payload.nurseName, completedAt]);

  const row = sheet.getLastRow();
  sheet.getRange(row, 1).setNumberFormat("d mmm yy");
  sheet.getRange(row, 4).setNumberFormat("yyyy-mm-dd hh:mm:ss");

  return { ok: true };
}

function getShiftInfo_() {
  const tz = Session.getScriptTimeZone() || "Asia/Bangkok";
  const now = new Date();
  const time = Number(Utilities.formatDate(now, tz, "HH")) * 60 +
               Number(Utilities.formatDate(now, tz, "mm"));

  const morningStart = 7 * 60 + 30;
  const afternoonStart = 15 * 60 + 30;
  const nightStart = 23 * 60 + 30;

  let shift;
  let dayOffset = 0;

  if (time >= morningStart && time < afternoonStart) {
    shift = "เช้า";
  } else if (time >= afternoonStart && time < nightStart) {
    shift = "บ่าย";
  } else {
    shift = "ดึก";
    if (time < morningStart) dayOffset = -1;
  }

  const roundDate = new Date(now);
  roundDate.setDate(roundDate.getDate() + dayOffset);

  const date = Utilities.formatDate(roundDate, tz, "yyyy-MM-dd");
  const day = Utilities.formatDate(roundDate, tz, "dd");
  const month = Utilities.formatDate(roundDate, tz, "MM");
  const year = Number(Utilities.formatDate(roundDate, tz, "yyyy")) + 543;

  return {
    date,
    displayDate: `${day}/${month}/${year}`,
    shift,
  };
}
