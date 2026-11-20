/****************************************************
 * 設備 QR Code 借還管理系統 API 後端（含 Token 驗證）
 ****************************************************/

const CONFIG = {
  EQUIPMENT_SHEET: "Equipment_List",
  RECORD_SHEET: "Borrow_Records",
  USER_SHEET: "Users",
  AVAILABLE: "可借用",
  BORROWED: "借出中",
  TIMEZONE: "Asia/Taipei",
  LOCK_TIMEOUT: 15000,
  HISTORY_LIMIT: 500,
  SESSION_TTL: 21600,      // 登入有效秒數（6 小時，CacheService 上限）
  MAX_LOGIN_FAILS: 5,      // 連續密碼錯誤上限
  FAIL_WINDOW: 600         // 鎖定秒數（10 分鐘）
};

/* ===================== API 入口 ===================== */

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    let result;

    if (action === "login") {
      result = handleLogin(contents.userId, contents.password);
    } else {
      // 除登入外，所有操作都必須通過 token 驗證
      const user = requireUser_(contents.token);

      switch (action) {
        case "getInitialData":
          result = getInitialData();
          break;
        case "borrow":
          result = processEquipment_("borrow", contents.id, user.name);
          break;
        case "return":
          result = processEquipment_("return", contents.id, user.name);
          break;
        case "getHistory":
          result = getBorrowHistory(contents.filter || "all");
          break;
        case "logout":
          CacheService.getScriptCache().remove("session:" + contents.token);
          result = { success: true, message: "已登出" };
          break;
        default:
          throw new Error("未知的操作類型：" + action);
      }
    }

    return createJsonResponse_(result);

  } catch (err) {
    return createJsonResponse_({
      success: false,
      code: err.code || "",
      message: err.message || "伺服器處理失敗。"
    });
  }
}

function createJsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("無法取得試算表。");
  return ss;
}

/* ===================== Session / Token ===================== */

function createSession_(user) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put("session:" + token, JSON.stringify(user), CONFIG.SESSION_TTL);
  return token;
}

function requireUser_(token) {
  const raw = token ? CacheService.getScriptCache().get("session:" + String(token)) : null;
  if (!raw) {
    const e = new Error("登入已過期，請重新登入。");
    e.code = "AUTH";
    throw e;
  }
  return JSON.parse(raw);
}

/* ===================== 登入 ===================== */

function handleLogin(userId, password) {
  const cleanUserId = String(userId || "").trim();
  const cleanPassword = String(password || "").trim();

  if (!cleanUserId) return { success: false, message: "請輸入帳號。" };
  if (!cleanPassword) return { success: false, message: "請輸入密碼。" };

  // 錯誤次數限制（避免暴力破解 4 位數密碼）
  const cache = CacheService.getScriptCache();
  const failKey = "fail:" + cleanUserId.slice(0, 100);
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= CONFIG.MAX_LOGIN_FAILS) {
    return { success: false, message: "嘗試次數過多，請 10 分鐘後再試。" };
  }

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.USER_SHEET);
  if (!sheet) {
    return { success: false, message: "找不到 " + CONFIG.USER_SHEET + " 使用者資料表。" };
  }

  const data = sheet.getDataRange().getDisplayValues();
  // 欄位順序: [0]user_id, [1]name, [2]role, [3]status, [4]password
  for (let i = 1; i < data.length; i++) {
    const rowUserId = String(data[i][0]).trim();
    const rowName   = String(data[i][1]).trim();
    const rowRole   = String(data[i][2]).trim();
    const rowStatus = String(data[i][3]).trim();
    const rowPass   = String(data[i][4]).trim();

    if (rowUserId === cleanUserId) {
      if (rowStatus === "不啟用") {
        return { success: false, message: "此帳號目前為「不啟用」狀態，無法登入。" };
      }
      const normalizedRowPass = rowPass === "0" ? "0000" : rowPass.padStart(4, "0");
      if (rowPass === cleanPassword || normalizedRowPass === cleanPassword) {
        cache.remove(failKey);
        const baseUser = { user_id: rowUserId, name: rowName, role: rowRole };
        const token = createSession_(baseUser);
        return {
          success: true,
          message: "登入成功",
          user: { user_id: rowUserId, name: rowName, role: rowRole, token: token }
        };
      } else {
        cache.put(failKey, String(fails + 1), CONFIG.FAIL_WINDOW);
        return { success: false, message: "密碼錯誤，請重新輸入。" };
      }
    }
  }

  cache.put(failKey, String(fails + 1), CONFIG.FAIL_WINDOW);
  return { success: false, message: "找不到此帳號（user_id）。" };
}

/* ===================== 工作表 ===================== */

function getEquipmentSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.EQUIPMENT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.EQUIPMENT_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([["設備編號", "設備名稱", "設備狀態", "QR Code", "廠牌"]]);
    formatHeader_(sheet, 5);
  }
  return sheet;
}

function getRecordSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.RECORD_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.RECORD_SHEET);
  }
  // 只在工作表為空時初始化標題與日期格式，避免每次呼叫都重設整欄格式
  if (sheet.getLastRow() === 0) {
    const headers = ["設備編號", "設備名稱", "廠牌", "借用人", "借用時間", "歸還人", "歸還時間"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader_(sheet, headers.length);
    sheet.getRange("E:E").setNumberFormat("yyyy/MM/dd HH:mm:ss");
    sheet.getRange("G:G").setNumberFormat("yyyy/MM/dd HH:mm:ss");
  }
  return sheet;
}

function formatHeader_(sheet, columns) {
  sheet.getRange(1, 1, 1, columns).setFontWeight("bold").setBackground("#dbeafe").setFontColor("#1e40af");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, columns);
}

/* ===================== 驗證與查詢 ===================== */

function validateInput_(id, name) {
  const cleanId = String(id || "").trim();
  const cleanName = String(name || "").trim();
  if (!cleanId) throw new Error("請輸入或掃描設備編號。");
  if (!cleanName) throw new Error("無法取得操作人姓名，請重新登入。");
  if (cleanName.length > 50) throw new Error("姓名不可超過 50 個字。");
  if (cleanId.length > 100) throw new Error("設備編號長度不正確。");
  return { id: cleanId, name: cleanName };
}

function findEquipment_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("設備清單目前沒有資料。");
  const values = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === id) {
      return {
        row: i + 2,
        id: String(values[i][0]).trim(),
        equipmentName: values[i][1],
        status: values[i][2],
        qrCode: values[i][3],
        brand: values[i][4]
      };
    }
  }
  throw new Error("找不到設備編號：" + id);
}

function getInitialData() {
  const ss = getSpreadsheet_();
  const equipmentSheet = getEquipmentSheet_(ss);
  const recordSheet = getRecordSheet_(ss);
  const equipment = [];

  if (equipmentSheet.getLastRow() >= 2) {
    const values = equipmentSheet.getRange(2, 1, equipmentSheet.getLastRow() - 1, 5).getDisplayValues();
    values.forEach(row => {
      if (String(row[0]).trim()) {
        equipment.push({
          id: String(row[0]).trim(),
          equipmentName: row[1],
          status: row[2],
          qrCode: row[3],
          brand: row[4]
        });
      }
    });
  }

  return {
    success: true,
    equipment: equipment,
    records: readHistory_(recordSheet)
  };
}

/* ===================== 借用 / 歸還 ===================== */

function processEquipment_(action, id, name) {
  const input = validateInput_(id, name);
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT);
    const ss = getSpreadsheet_();
    const equipmentSheet = getEquipmentSheet_(ss);
    const recordSheet = getRecordSheet_(ss);
    const equipment = findEquipment_(equipmentSheet, input.id);
    const now = new Date();

    if (action === "borrow") {
      if (equipment.status !== CONFIG.AVAILABLE) {
        throw new Error("此設備目前為「" + equipment.status + "」，無法借用。");
      }
      let appendedRow = 0;
      try {
        recordSheet.appendRow([equipment.id, equipment.equipmentName, equipment.brand, input.name, now, "", ""]);
        appendedRow = recordSheet.getLastRow();
        equipmentSheet.getRange(equipment.row, 3).setValue(CONFIG.BORROWED);
      } catch (err) {
        if (appendedRow) recordSheet.deleteRow(appendedRow);
        throw err;
      }
    } else if (action === "return") {
      if (equipment.status !== CONFIG.BORROWED) {
        throw new Error("此設備目前不是借出狀態，無法歸還。");
      }
      const lastRow = recordSheet.getLastRow();
      if (lastRow < 2) throw new Error("找不到尚未歸還的借用紀錄。");

      const rows = recordSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      let targetRow = -1;

      for (let i = rows.length - 1; i >= 0; i--) {
        const rowId = String(rows[i][0]).trim();
        const returnTime = rows[i][6];
        if (rowId === equipment.id && (!returnTime || returnTime === "")) {
          targetRow = i + 2;
          break;
        }
      }

      if (targetRow < 0) throw new Error("找不到此設備尚未歸還的借用紀錄。");

      const oldValues = recordSheet.getRange(targetRow, 6, 1, 2).getValues();
      try {
        recordSheet.getRange(targetRow, 6, 1, 2).setValues([[input.name, now]]);
        equipmentSheet.getRange(equipment.row, 3).setValue(CONFIG.AVAILABLE);
      } catch (err) {
        recordSheet.getRange(targetRow, 6, 1, 2).setValues(oldValues);
        throw err;
      }
    } else {
      throw new Error("未知的操作類型。");
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      action: action,
      id: equipment.id,
      equipmentName: equipment.equipmentName,
      brand: equipment.brand,
      person: input.name,
      time: formatTime_(now),
      status: action === "borrow" ? CONFIG.BORROWED : CONFIG.AVAILABLE,
      message: action === "borrow" ? "設備借用成功！" : "設備歸還成功！"
    };

  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/* ===================== 歷史紀錄 ===================== */

function getBorrowHistory(filter) {
  const sheet = getRecordSheet_(getSpreadsheet_());
  let records = readHistory_(sheet);
  if (filter === "outstanding") {
    records = records.filter(record => !record.returnTime);
  }
  return { success: true, records: records };
}

function readHistory_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const startRow = Math.max(2, lastRow - CONFIG.HISTORY_LIMIT + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues();

  return values.reverse().map(row => ({
    id: String(row[0] || ""),
    equipmentName: row[1] || "",
    brand: row[2] || "",
    borrower: row[3] || "",
    borrowTime: formatTime_(row[4]),
    returner: row[5] || "",
    returnTime: formatTime_(row[6])
  }));
}

function formatTime_(value) {
  if (!value) return "";
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyy/MM/dd HH:mm:ss");
}

function doGet() {
  return ContentService.createTextOutput("API 新版（含 token 驗證）運作中");
}
