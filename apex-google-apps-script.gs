/**
 * ════════════════════════════════════════════════════════
 * APEX FITNESS — Google Apps Script v4.2
 * ✅ ระบบ Walk-in / รายครั้ง จ่าย QR + อัปโหลดสลิป
 * ✅ กฎยกเลิกก่อน X ชม. → No-show อัตโนมัติ
 * ✅ เปิด/ปิดรับจองล่วงหน้า (booking window)
 * ✅ รูปโปรไฟล์สมาชิก (เก็บใน Google Drive)
 * ✅ LINE LIFF Login + Messaging API Push (ใส่ค่าเองได้ทีหลัง)
 * ✅ แจ้งเตือนหมดอายุ + โปรโมชั่นต่ออายุ
 * ✅ LockService ป้องกันจองซ้อนตอนใช้งานพร้อมกันหลายคน
 * ✅ [v4.1] แก้ Drive URL → lh3.googleusercontent.com (embed ใน <img> ได้)
 * ✅ [v4.1] handleSaveSettings auto-upload base64 → Drive URL (ป้องกัน cell overflow)
 * ✅ [v4.1] uploadImage action ใหม่ สำหรับอัปโหลดรูปตาราง / QR
 * ✅ [v4.2] exportReport → สร้าง Google Spreadsheet ใหม่ return URL (members/bookings/attendance/guestBookings)
 * ✅ [v4.2] notifyAdmin push LINE ไปหา ADMIN_LINE_USER_IDS จริงๆ (ไม่ใช่แค่ Logger.log)
 * ✅ [v4.2] getSessionBookings → return รายชื่อสมาชิก + guest ที่จองคลาสนั้น
 * ⚠️ LINE Notify (notify-api.line.me) ปิดให้บริการแล้วตั้งแต่ 31 มี.ค. 2025
 *     ใช้ LINE Messaging API แทน — ตั้ง ADMIN_LINE_USER_IDS ใน CONFIG
 * ════════════════════════════════════════════════════════
 */

// ════════════════════════════════════════════════════════
// ⚙️  CONFIG — ใส่ค่าที่ได้จาก LINE Developers Console ตรงนี้
// ════════════════════════════════════════════════════════
var CONFIG = {
  ADMIN_PIN:       "apex2024",
  STAFF_PIN:       "staff2024",
  NOSHOW_LIMIT:    3,
  SUSPENSION_DAYS: 14,
  // ── LINE Messaging API (ใหม่ ใช้แทน LINE Notify ที่ปิดไปแล้ว) ──
  LIFF_ID:"2010458588-pD1CwMsn",
  LINE_CHANNEL_ACCESS_TOKEN:"PYJ4WfoJfEMOGu2o2ZbcKsAfuQ2UYGHO5PdnUOKGBZ9JFGZ05ntAMYfJtTt3TnDI2AvntdijdPHybZqRknYTnryTWgMvz2B9MJhvt2cySpbjRsliv6QIUYmlJnwdHMndRFXjZDNqWmFXfiu00N1hpwdB04t89/1O/w1cDnyilFU=",
  // ✅ [v4.2] ใส่ lineUserId ของแอดมิน/เจ้าของที่ต้องการรับแจ้งเตือน
  // วิธีหา lineUserId: ให้แอดมิน add LINE OA เป็นเพื่อน แล้ว echo lineUserId กลับมาผ่าน webhook
  // หรือดูได้จาก LINE Developers Console → Messaging API → Webhook events
  ADMIN_LINE_USER_IDS: [],   // เช่น ["Ufa2e785ec3540c692fc694f953e4cfd3", "Uyyy..."]
  DRIVE_FOLDER_ID: "",
};

var SHEETS = {
  MEMBERS:        "Members",
  CLASSES:        "Classes",
  SESSIONS:       "Sessions",
  BOOKINGS:       "Bookings",
  ATTENDANCE:     "Attendance",
  NOTIFICATIONS:  "Notifications",
  SETTINGS:       "Settings",
  APPDATA:        "AppData",
  GUEST_BOOKINGS: "GuestBookings",
  FEEDBACK:       "Feedback",
  STAFF:          "Staff",
  ETIQUETTE:      "Etiquette",
};

// ════════════════════════════════════════════════════════
// 📁 GOOGLE DRIVE — เก็บรูปสลิป / รูปโปรไฟล์
// ════════════════════════════════════════════════════════
function getUploadFolder() {
  if (CONFIG.DRIVE_FOLDER_ID) {
    try { return DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID); } catch(e) {}
  }
  var name = "Apex Fitness Uploads";
  var it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

// ✅ [v4.1 FIX] เปลี่ยน URL จาก uc?export=view (deprecated/ไม่ embed ได้)
//              → lh3.googleusercontent.com/d/FILE_ID ซึ่ง embed ใน <img> ได้โดยตรง
function saveBase64Image(base64Data, fileName) {
  try {
    if (!base64Data) return "";
    // ตรวจ MIME type จาก data URL prefix
    var mimeMatch = base64Data.match(/^data:(image\/\w+);base64,/);
    var mimeType  = mimeMatch ? mimeMatch[1] : "image/jpeg";
    var clean  = base64Data.replace(/^data:image\/\w+;base64,/, "");
    var bytes  = Utilities.base64Decode(clean);
    var blob   = Utilities.newBlob(bytes, mimeType, fileName);
    var folder = getUploadFolder();
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // lh3 URL: embed ได้โดยตรงใน <img src="..."> ไม่ต้อง redirect
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch(e) {
    Logger.log("saveBase64Image error: " + e);
    return "";
  }
}

// ════════════════════════════════════════════════════════
// 🔔 LINE MESSAGING API (ใหม่) — Push ข้อความหาลูกค้ารายคน
// ════════════════════════════════════════════════════════
function pushLineMessage(lineUserId, text) {
  if (!lineUserId || !CONFIG.LINE_CHANNEL_ACCESS_TOKEN) return false;
  try {
    var resp = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + CONFIG.LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify({ to: lineUserId, messages: [{ type: "text", text: text }] }),
      muteHttpExceptions: true,
    });
    var ok = resp.getResponseCode() === 200;
    Logger.log("LINE push " + (ok ? "OK" : "FAIL " + resp.getContentText()));
    return ok;
  } catch(e) { Logger.log("pushLineMessage error: " + e); return false; }
}

// ⚠️ เก่า — LINE Notify ปิดบริการแล้ว (31 มี.ค. 2025)
function sendLineNotify(token, message) { return false; }

// ✅ [v4.2] แจ้งเตือนแอดมินผ่าน LINE Messaging API
// ต้องใส่ lineUserId ของแอดมินใน CONFIG.ADMIN_LINE_USER_IDS
function notifyAdmin(msg) {
  Logger.log("[AdminLog] " + msg);
  var ids = CONFIG.ADMIN_LINE_USER_IDS || [];
  ids.forEach(function(uid) { pushLineMessage(uid, "🔔 [Admin] Apex Fitness\n" + msg); });
}
function notifyOwner(msg) {
  Logger.log("[OwnerLog] " + msg);
  var ids = CONFIG.ADMIN_LINE_USER_IDS || [];
  ids.forEach(function(uid) { pushLineMessage(uid, "👑 [Owner] Apex Fitness\n" + msg); });
}

// ════════════════════════════════════════════════════════
// 🔐 LINE LOGIN (LIFF) — ผูก lineUserId เข้ากับสมาชิก
// ════════════════════════════════════════════════════════
function handleLineLogin(d) {
  if (!d.lineUserId) return { ok:false, error:"missing lineUserId" };
  var sh = getSheet(SHEETS.MEMBERS), data = sh.getDataRange().getValues();
  var headers = data[0];
  var lineCol = headers.indexOf("lineUserId");
  if (lineCol === -1) return { ok:false, error:"lineUserId column not found — รัน migrateAddColumns() ก่อน" };

  for (var i=1; i<data.length; i++) {
    if (data[i][lineCol] === d.lineUserId) {
      return { ok:true, type:"user", member: rowToObj(headers, data[i]) };
    }
  }
  if (d.phone) {
    var cp = String(d.phone).replace(/\D/g,"");
    for (var j=1; j<data.length; j++) {
      var mp = String(data[j][headers.indexOf("phone")] || "").replace(/\D/g,"");
      if (mp===cp || "0"+mp===cp || mp==="0"+cp) {
        sh.getRange(j+1, lineCol+1).setValue(d.lineUserId);
        var photoCol = headers.indexOf("photoUrl");
        if (photoCol > -1 && d.pictureUrl && !data[j][photoCol]) sh.getRange(j+1, photoCol+1).setValue(d.pictureUrl);
        data[j][lineCol] = d.lineUserId;
        return { ok:true, type:"user", member: rowToObj(headers, data[j]) };
      }
    }
  }
  return { ok:false, needPhone:true, error:"ไม่พบบัญชีที่ผูกกับ LINE นี้ กรุณากรอกเบอร์โทรเพื่อเชื่อมบัญชี" };
}

function rowToObj(headers, row) {
  var obj = {};
  headers.forEach(function(h,i){ obj[h] = row[i]; });
  return obj;
}

// ════════════════════════════════════════════════════════
// 📊 DAILY SUMMARY
// ════════════════════════════════════════════════════════
function sendDailyAdminSummary() {
  var today    = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var sessions = readSheet(SHEETS.SESSIONS).filter(function(s){ return s.id && String(s.id).indexOf(today) !== -1; });
  var members  = readSheet(SHEETS.MEMBERS);
  var totalBk  = 0, totalWL = 0;
  sessions.forEach(function(s){
    totalBk += String(s.bookings||"").split(",").filter(Boolean).length;
    totalWL += String(s.waitlist||"").split(",").filter(Boolean).length;
  });
  var pendingGuests = readSheet(SHEETS.GUEST_BOOKINGS).filter(function(g){ return g.status === "pending"; }).length;
  var msg = [
    "📊 สรุปประจำวัน Apex Fitness — " + today,
    "📅 คลาสวันนี้: " + sessions.length,
    "🎫 จองทั้งหมด: " + totalBk + " | 🪑 สำรอง: " + totalWL,
    "👥 สมาชิก: " + members.length,
    "💳 Walk-in รอตรวจสลิป: " + pendingGuests,
  ].join("\n");
  Logger.log(msg);
  // ✅ [v4.2] Push LINE ไปหาแอดมินทุกคนที่ลงทะเบียนไว้
  var ids = CONFIG.ADMIN_LINE_USER_IDS || [];
  ids.forEach(function(uid) { pushLineMessage(uid, msg); });
}

// ════════════════════════════════════════════════════════
// ⚠️ แจ้งหมดอายุ + โปรโมชั่นต่ออายุ (trigger ทุกวัน 09:00)
// ════════════════════════════════════════════════════════
function sendExpiryReminders() {
  var settings = readSettingsObj();
  var reminderDays = [parseInt(settings.expiryReminderDays) || 5, 3, 1];
  var promo1 = settings.renewalPromo1Month || "100";
  var promo3 = settings.renewalPromo3Month || "300";
  var members = readSheet(SHEETS.MEMBERS);

  members.forEach(function(m){
    if (normalizeMembershipType(m.membershipType) !== "monthly" || !m.membershipExpiry) return;
    var diff = Math.ceil((new Date(m.membershipExpiry) - new Date()) / 86400000);
    if (reminderDays.indexOf(diff) === -1) return;

    var body = "สมาชิกของคุณจะหมดอายุในอีก " + diff + " วัน (" + m.membershipExpiry + ")\n\n" +
      "🎁 โปรต่ออายุตอนนี้:\n" +
      "• ต่อ 1 เดือน ลด " + promo1 + " บาท\n" +
      "• ต่อ 3 เดือน ลด " + promo3 + " บาท\n\n" +
      "ติดต่อต่ออายุได้ที่ฟิตเนสหรือ LINE OA";

    notify(m.id, "membership", "⏰ สมาชิกใกล้หมดอายุ + โปรต่ออายุ", body);
    if (m.lineUserId) pushLineMessage(m.lineUserId, "🔔 Apex Fitness\n" + body);
  });
}

// ════════════════════════════════════════════════════════
// 🎉 แจ้ง Waitlist ได้ที่นั่ง
// ════════════════════════════════════════════════════════
function notifyWaitlistPromotion(memberId, sessionId, className) {
  var member = readSheet(SHEETS.MEMBERS).filter(function(m){ return m.id === memberId; })[0];
  if (!member) return;
  var body = "มีที่ว่างใน " + className + " — เลื่อนจากสำรองเป็นจองหลักแล้ว เจอกันนะ!";
  notify(memberId, "promoted", "🎉 คุณได้ที่นั่งแล้ว!", body);
  if (member.lineUserId) pushLineMessage(member.lineUserId, "🎉 Apex Fitness\n" + body);
}

function notifySuspension(memberId, noShowCount, suspendedUntil) {
  var member = readSheet(SHEETS.MEMBERS).filter(function(m){ return m.id === memberId; })[0];
  if (!member) return;
  var body = "No-show ครบ " + noShowCount + " ครั้ง ระงับการจองถึง " + suspendedUntil;
  notify(memberId, "suspended", "🚫 บัญชีถูกระงับ", body);
  if (member.lineUserId) pushLineMessage(member.lineUserId, "⚠️ Apex Fitness\n" + body);
}

// ════════════════════════════════════════════════════════
// 🔧 SETUP SHEETS (รันครั้งแรก)
// ════════════════════════════════════════════════════════
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = {};
  defs[SHEETS.MEMBERS]        = ["id","name","phone","pin","lineId","membershipType","membershipExpiry","sessionsLeft","noShowCount","suspended","suspendedUntil","deviceId","joinDate","photoUrl","lineUserId"];
  defs[SHEETS.CLASSES]        = ["id","name","instructor","time","duration","capacity","waitlistCap","days","color","icon","popular","active"];
  defs[SHEETS.SESSIONS]       = ["id","classId","date","bookings","waitlist"];
  defs[SHEETS.BOOKINGS]       = ["id","sessionId","memberId","type","createdAt"];
  defs[SHEETS.ATTENDANCE]     = ["sessionId","memberId","status","markedAt"];
  defs[SHEETS.NOTIFICATIONS]  = ["id","memberId","type","title","body","timestamp","read"];
  defs[SHEETS.SETTINGS]       = ["key","value"];
  defs[SHEETS.APPDATA]        = ["key","value"];
  defs[SHEETS.GUEST_BOOKINGS] = ["id","name","phone","sessionId","className","classDate","classTime","amount","slipUrl","status","createdAt","confirmedBy","confirmedAt"];
  defs[SHEETS.FEEDBACK]       = ["id","memberId","memberName","type","message","status","createdAt","response","respondedAt"];

  Object.keys(defs).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    var h = defs[name];
    sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight("bold").setBackground("#F97316").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
  });

  var settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (settingsSheet.getLastRow() < 2) {
    settingsSheet.getRange(2,1,15,2).setValues([
      ["contactPhone","084-220-9391"],
      ["contactLine","@478xvozx"],
      ["lineOfficialUrl","https://line.me/R/ti/p/@478xvozx"],
      ["noShowLimit","3"],
      ["suspensionDays","14"],
      ["checkinBeforeMin","10"],
      ["checkinAfterMin","5"],
      ["gymName","Apex Fitness Chiang Mai"],
      ["cancelHoursLimit","2"],
      ["bookingOpenDaysAhead","7"],
      ["bookingCloseHoursBefore","2"],
      ["expiryReminderDays","5"],
      ["renewalPromo1Month","100"],
      ["renewalPromo3Month","300"],
      ["guestPrice","150"],
    ]);
  }
  Logger.log("✅ setupSheets complete!");
  try { SpreadsheetApp.getUi().alert("✅ Setup สำเร็จ! ทุก Sheet พร้อมใช้งาน"); } catch(e) {}
}

// ════════════════════════════════════════════════════════
// 🔄 MIGRATE — เพิ่มคอลัมน์ใหม่ให้ Sheet เดิม (รันครั้งเดียว)
// ════════════════════════════════════════════════════════
function migrateAddColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var msh = ss.getSheetByName(SHEETS.MEMBERS);
  var mh = msh.getRange(1,1,1,msh.getLastColumn()).getValues()[0];
  var added = [];
  if (mh.indexOf("photoUrl") === -1) { msh.getRange(1, mh.length+1).setValue("photoUrl"); added.push("photoUrl"); mh.push("photoUrl"); }
  if (mh.indexOf("lineUserId") === -1) { msh.getRange(1, mh.length+1).setValue("lineUserId"); added.push("lineUserId"); }

  if (!ss.getSheetByName(SHEETS.GUEST_BOOKINGS)) {
    var gb = ss.insertSheet(SHEETS.GUEST_BOOKINGS);
    var h = ["id","name","phone","sessionId","className","classDate","classTime","amount","slipUrl","status","createdAt","confirmedBy","confirmedAt"];
    gb.getRange(1,1,1,h.length).setValues([h]).setFontWeight("bold").setBackground("#F97316").setFontColor("#FFFFFF");
    gb.setFrozenRows(1);
    added.push("GuestBookings sheet");
  }

  if (!ss.getSheetByName(SHEETS.FEEDBACK)) {
    var fb = ss.insertSheet(SHEETS.FEEDBACK);
    var fh = ["id","memberId","memberName","type","message","status","createdAt","response","respondedAt"];
    fb.getRange(1,1,1,fh.length).setValues([fh]).setFontWeight("bold").setBackground("#F97316").setFontColor("#FFFFFF");
    fb.setFrozenRows(1);
    added.push("Feedback sheet");
  }

  var ssh = ss.getSheetByName(SHEETS.SETTINGS);
  var sdata = ssh.getDataRange().getValues();
  var existingKeys = sdata.slice(1).map(function(r){ return r[0]; });
  var defaults = [
    ["cancelHoursLimit","2"], ["bookingOpenDaysAhead","7"], ["bookingCloseHoursBefore","2"],
    ["expiryReminderDays","5"], ["renewalPromo1Month","100"], ["renewalPromo3Month","300"], ["guestPrice","150"],
    ["scheduleImage",""], ["maxActiveBookings","5"],
  ];
  defaults.forEach(function(pair){
    if (existingKeys.indexOf(pair[0]) === -1) { ssh.appendRow(pair); added.push(pair[0]); }
  });

  var msg = added.length ? ("✅ เพิ่มแล้ว: " + added.join(", ")) : "ไม่มีอะไรต้องเพิ่ม (อัปเดตล่าสุดแล้ว)";
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// ════════════════════════════════════════════════════════
// 🔄 FIX CLASS SCHEDULE
// ════════════════════════════════════════════════════════
function fixClassSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.CLASSES);
  if (!sh) { Logger.log("Classes sheet not found"); return; }
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 12).clearContent();

  var classes = [
    ["c1",  "Yoga Hatha",         "ครูแอร์",   "9:00",  60, 12, 8, "1",   "#F97316", "🧘",  false, true],
    ["c3",  "Fit Ball",           "ครูนาย",    "10:10", 60, 12, 8, "2",   "#EF4444", "🥊",  false, true],
    ["c6",  "Yoga Gentle Flow",   "ครูดีใจ",   "9:00",  60, 12, 8, "3",   "#14B8A6", "🧘",  false, true],
    ["c8",  "Pilates Mat",        "ครูเมรี่",  "10:10", 60, 15, 8, "4",   "#10B981", "⚡",  false, true],
    ["c9",  "Yoga Calm Alignment","ครูใหม่",   "10:10", 60, 12, 8, "5",   "#F97316", "🧘",  false, true],
    ["c11", "Yoga Power",         "ครูแอร์",   "11:00", 60, 12, 8, "6",   "#F97316", "💪",  false, true],
    ["c2",  "Step Dance",         "ครูนาย",    "19:10", 60, 12, 8, "1",   "#22C55E", "💃",  false, true],
    ["c4",  "Zumba Dance",        "ครูขวัญ",   "18:00", 60, 15, 8, "2",   "#A855F7", "💃",  true,  true],
    ["c5",  "Body Shape Up",      "ครูเภา",    "19:10", 60, 20, 8, "2",   "#F59E0B", "🏋️", true,  true],
    ["c7",  "Functional Training","ครูอาร์ม",  "18:30", 60, 20, 8, "3",   "#10B981", "🏋️", true,  true],
    ["c12", "Zumba Dance",        "ครูจ๊ะโอ๋", "18:00", 60, 15, 8, "4",   "#A855F7", "💃",  true,  true],
    ["c10", "Yoga Wheel",         "ครูแอร์",   "19:10", 60, 12, 8, "4",   "#F97316", "🧘",  true,  true],
    ["c13", "Body Shape Up",      "ครูเภา",    "18:30", 60, 20, 8, "5",   "#F59E0B", "🏋️", true,  true],
  ];
  sh.getRange(2, 1, classes.length, 12).setValues(classes);
  sh.getRange(2, 4, classes.length, 1).setNumberFormat("@STRING@");
  Logger.log("✅ Class schedule updated: " + classes.length + " classes");
  try { SpreadsheetApp.getUi().alert("✅ อัปเดตตารางคลาสสำเร็จ! " + classes.length + " คลาส"); } catch(e) {}
}

// ════════════════════════════════════════════════════════
// 📅 สร้าง Sessions 14 วัน
// ════════════════════════════════════════════════════════
function generateSessionsForNext14Days() {
  var classes = readSheet(SHEETS.CLASSES).filter(function(c){
    return c.active === true || String(c.active).toUpperCase() === "TRUE";
  });
  var sh       = getSheet(SHEETS.SESSIONS);
  var existing = readSheet(SHEETS.SESSIONS).map(function(s){ return String(s.id); });
  var today    = new Date();
  var added    = 0;

  for (var d = 0; d < 14; d++) {
    var date = new Date(today);
    date.setDate(date.getDate() + d);
    var ds  = Utilities.formatDate(date, "GMT+7", "yyyy-MM-dd");
    var dow = date.getDay();
    classes.forEach(function(c) {
      var days = String(c.days || "").split(",").map(function(x){ return parseInt(x.trim()); }).filter(function(n){ return !isNaN(n); });
      if (days.indexOf(dow) === -1) return;
      var sid = c.id + "_" + ds;
      if (existing.indexOf(sid) !== -1) return;
      sh.appendRow([sid, c.id, ds, "", ""]);
      existing.push(sid);
      added++;
    });
  }
  Logger.log("✅ Sessions generated: " + added + " new sessions");
}

// ════════════════════════════════════════════════════════
// 🌐 HTTP HANDLERS
// ════════════════════════════════════════════════════════
function doGet(e) {
  var action = (e.parameter && e.parameter.action) || "getAll";
  try {
    if (action === "getAll")           return json(getAllData());
    if (action === "getData")          return json(handleGetData(e.parameter.key));
    if (action === "login")            return json(handleLogin(e.parameter.phone, e.parameter.pin, e.parameter.deviceId));
    if (action === "getSessionBookings") return json(handleGetSessionBookings(e.parameter.sessionId));  // ✅ [v4.2]
    return json({ error: "Unknown GET action: " + action });
  } catch(err) { return json({ error: err.toString() }); }
}

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action;
    if (action === "book")                  return json(handleBook(data));
    if (action === "joinWaitlist")          return json(handleJoinWaitlist(data));
    if (action === "cancel")                return json(handleCancel(data));
    if (action === "cancelWaitlist")        return json(handleCancelWaitlist(data));
    if (action === "checkin")               return json(handleCheckin(data));
    if (action === "markAttendance")        return json(handleMarkAttendance(data));
    if (action === "promoteWaitlist")       return json(handlePromoteWaitlist(data));
    if (action === "saveMember")            return json(handleSaveMember(data));
    if (action === "unlockDevice")          return json(handleUnlockDevice(data));
    if (action === "saveSettings")          return json(handleSaveSettings(data));
    if (action === "saveEtiquette")         return json(handleSaveEtiquette(data));
    if (action === "setData")               return json(handleSetData(data.key, data.value));
    if (action === "syncMembersFromAristo") return json(handleSyncAristo(data));
    if (action === "sendAdminSummary")      { sendDailyAdminSummary(); return json({ok:true}); }
    if (action === "lineLogin")             return json(handleLineLogin(data));
    if (action === "submitGuestBooking")    return json(handleSubmitGuestBooking(data));
    if (action === "confirmGuestPayment")   return json(handleConfirmGuestPayment(data));
    if (action === "uploadMemberPhoto")     return json(handleUploadMemberPhoto(data));
    if (action === "uploadImage")           return json(handleUploadImage(data));   // ✅ [v4.1 NEW]
    if (action === "exportReport")          return json(handleExportReport(data));  // ✅ [v4.2 NEW]
    if (action === "submitFeedback")        return json(handleSubmitFeedback(data));
    if (action === "resolveFeedback")       return json(handleResolveFeedback(data));
    if (action === "saveStaff")             return json(handleSaveStaff(data));
    if (action === "deleteStaff")           return json(handleDeleteStaff(data));
    if (action === "changeAdminPin")        return json(handleChangeAdminPin(data));
    return json({ error: "Unknown POST action: " + action });
  } catch(err) { return json({ error: err.toString() }); }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════
// 📋 DATA HELPERS
// ════════════════════════════════════════════════════════
function getSheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight("bold"); }
  return sh;
}
function readSheet(name) {
  var sh = getSheet(name);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).filter(function(row){ return row[0] !== ""; }).map(function(row) {
    var obj = {}; headers.forEach(function(h,i){ obj[h] = row[i]; }); return obj;
  });
}
function readSettingsObj() {
  var obj = {};
  readSheet(SHEETS.SETTINGS).forEach(function(s){ if (s.key) obj[s.key] = s.value; });
  return obj;
}
function formatTimeValue(val) {
  if (!val && val !== 0) return "08:00";
  if (val instanceof Date) {
    var h = val.getHours(), m = val.getMinutes();
    return (h < 10 ? "0" : "") + h + ":" + (m < 10 ? "0" : "") + m;
  }
  var s = String(val);
  if (/^\d{1,2}:\d{2}/.test(s)) return s.substring(0, 5);
  var m2 = s.match(/T(\d{2}):(\d{2})/);
  if (m2) {
    var h2 = (parseInt(m2[1]) + 7) % 24, min = parseInt(m2[2]);
    return (h2 < 10 ? "0" : "") + h2 + ":" + (min < 10 ? "0" : "") + min;
  }
  return "08:00";
}
function getAllData() {
  var classesRaw = readSheet(SHEETS.CLASSES);
  var classes = classesRaw.map(function(c) { c.time = formatTimeValue(c.time); return c; });
  try { generateSessionsForNext14Days(); } catch(e) { Logger.log("generateSessions err: "+e); }
  return {
    members:       readSheet(SHEETS.MEMBERS),
    classes:       classes,
    sessions:      readSheet(SHEETS.SESSIONS),
    bookings:      readSheet(SHEETS.BOOKINGS),
    attendance:    readSheet(SHEETS.ATTENDANCE),
    notifications: readSheet(SHEETS.NOTIFICATIONS).map(function(n) {
      // ts may be stored as a Date cell → convert to ISO string
      if (n.ts instanceof Date) n.ts = n.ts.toISOString();
      return n;
    }),
    settings:      readSheet(SHEETS.SETTINGS),
    guestBookings: readSheet(SHEETS.GUEST_BOOKINGS).map(function(g) {
      // classTime is stored as a Time cell in Sheets → GAS returns Date(1899-12-30T...) → format it
      g.classTime = formatTimeValue(g.classTime);
      // createdAt/confirmedAt: ensure string
      if (g.createdAt instanceof Date) g.createdAt = g.createdAt.toISOString();
      if (g.confirmedAt instanceof Date) g.confirmedAt = g.confirmedAt.toISOString();
      return g;
    }),
    feedback:      readSheet(SHEETS.FEEDBACK),
    staff:         readStaffList(),
    etiquetteRules: readEtiquetteRules(),
  };
}

// ── [v4.3] Staff Management ──────────────────────────────────────────────────
function readStaffList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.STAFF)) return [];
  return readSheet(SHEETS.STAFF).map(function(s) {
    return { id: String(s.id||""), name: String(s.name||""), pin: String(s.pin||""), active: s.active === true || String(s.active).toUpperCase() === "TRUE", note: String(s.note||"") };
  }).filter(function(s) { return s.id && s.name; });
}

function getOrCreateStaffSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.STAFF);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.STAFF);
    sh.getRange(1,1,1,5).setValues([["id","name","pin","active","note"]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function handleSaveStaff(d) {
  if (!d.name || !d.pin) return { ok:false, error:"ต้องระบุชื่อและ PIN" };
  if (d.pin === CONFIG.ADMIN_PIN) return { ok:false, error:"PIN นี้ใช้โดย Admin อยู่แล้ว" };
  var sh = getOrCreateStaffSheet();
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf("id"), nameCol = headers.indexOf("name"), pinCol = headers.indexOf("pin"), activeCol = headers.indexOf("active"), noteCol = headers.indexOf("note");
  var staffId = d.id || "staff_" + new Date().getTime();
  // Check duplicate PIN (except own row)
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) !== String(staffId) && String(data[i][pinCol]) === String(d.pin) && (data[i][activeCol] === true || String(data[i][activeCol]).toUpperCase() === "TRUE")) {
      return { ok:false, error:"PIN นี้ถูกใช้โดย " + data[i][nameCol] + " อยู่แล้ว" };
    }
  }
  if (d.id) {
    // Update existing
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][idCol]) === String(d.id)) {
        sh.getRange(j+1, nameCol+1).setValue(d.name);
        sh.getRange(j+1, pinCol+1).setValue(d.pin);
        sh.getRange(j+1, activeCol+1).setValue(d.active !== false ? true : false);
        sh.getRange(j+1, noteCol+1).setValue(d.note || "");
        return { ok:true, id: d.id };
      }
    }
    return { ok:false, error:"ไม่พบ staff นี้" };
  } else {
    // Insert new
    sh.appendRow([staffId, d.name, d.pin, true, d.note || ""]);
    return { ok:true, id: staffId };
  }
}

function handleDeleteStaff(d) {
  if (!d.id) return { ok:false, error:"missing id" };
  var sh = getOrCreateStaffSheet();
  var data = sh.getDataRange().getValues();
  var idCol = data[0].indexOf("id");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(d.id)) {
      sh.deleteRow(i+1);
      return { ok:true };
    }
  }
  return { ok:false, error:"ไม่พบ staff นี้" };
}
// ── [Etiquette] กฎ Gym Etiquette (แก้ไขได้จาก Admin) ──────────────────────────
var DEFAULT_ETIQUETTE = [
  { id:"e1", thai:"อย่าเสียงดัง / คุยโทรศัพท์ระหว่างคลาส", eng:"No loud talking / phone calls during class" },
  { id:"e2", thai:"วางอุปกรณ์คืนที่เดิมหลังใช้งาน", eng:"Return equipment after use" },
  { id:"e3", thai:"มาตรงเวลา ไม่เข้าสาย", eng:"Be on time, no late entry" },
  { id:"e4", thai:"ไม่ใช้โทรศัพท์ระหว่างคลาส", eng:"No phone use / social media during class" },
  { id:"e5", thai:"แต่งกายเหมาะสม สะอาดเป็นระเบียบ", eng:"Wear appropriate, clean attire" },
  { id:"e6", thai:"เคารพเพื่อนสมาชิกและครูผู้สอน", eng:"Respect fellow members and instructors" },
];

function getOrCreateEtiquetteSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.ETIQUETTE);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.ETIQUETTE);
    sh.getRange(1,1,1,3).setValues([["id","thai","eng"]]);
    sh.setFrozenRows(1);
    DEFAULT_ETIQUETTE.forEach(function(r, i) {
      sh.getRange(i+2,1,1,3).setValues([[r.id, r.thai, r.eng]]);
    });
  }
  return sh;
}

function readEtiquetteRules() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(SHEETS.ETIQUETTE)) return DEFAULT_ETIQUETTE;
    var sh = ss.getSheetByName(SHEETS.ETIQUETTE);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return DEFAULT_ETIQUETTE;
    var headers = data[0];
    var idC = headers.indexOf("id"), thC = headers.indexOf("thai"), enC = headers.indexOf("eng");
    var rules = [];
    for (var i = 1; i < data.length; i++) {
      var thai = String(data[i][thC]||"").trim();
      if (!thai) continue;
      rules.push({ id: String(data[i][idC]||("e"+(i))), thai: thai, eng: String(data[i][enC]||"") });
    }
    return rules.length > 0 ? rules : DEFAULT_ETIQUETTE;
  } catch(e) { return DEFAULT_ETIQUETTE; }
}

function handleSaveEtiquette(d) {
  if (!d.rules || !Array.isArray(d.rules)) return { ok:false, error:"missing rules" };
  var sh = getOrCreateEtiquetteSheet();
  // Clear existing data rows
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow-1, 3).clearContent();
  // Write new rules
  d.rules.forEach(function(r, i) {
    var rId = String(r.id || ("e"+(i+1)));
    var rThai = String(r.thai || "").trim();
    var rEng = String(r.eng || "").trim();
    if (!rThai) return;
    sh.getRange(i+2,1,1,3).setValues([[rId, rThai, rEng]]);
  });
  return { ok:true };
}

function handleGetData(key) {
  if (!key) return { value: null };
  var sh = getOrCreateSheet(SHEETS.APPDATA, ["key","value"]);
  var data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) { if (String(data[i][0]) === String(key)) return { value: data[i][1] }; }
  return { value: null };
}
function handleSetData(key, value) {
  if (!key) return { ok: false };
  var sh = getOrCreateSheet(SHEETS.APPDATA, ["key","value"]);
  var data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) { if (String(data[i][0]) === String(key)) { sh.getRange(i+1,2).setValue(value); return { ok: true }; } }
  sh.appendRow([key, value]);
  return { ok: true };
}

// ════════════════════════════════════════════════════════
// 🔐 LOGIN
// ════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════
// 🔑 เปลี่ยน Admin PIN — เก็บใน PropertiesService (override CONFIG)
// ════════════════════════════════════════════════════════
function getAdminPin() {
  var override = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN_OVERRIDE");
  return override || CONFIG.ADMIN_PIN;
}

function handleChangeAdminPin(d) {
  if (!d.currentPin || !d.newPin) return { ok: false, error: "ข้อมูลไม่ครบ" };
  if (String(d.currentPin) !== getAdminPin()) return { ok: false, error: "รหัสเก่าไม่ถูกต้อง" };
  if (String(d.newPin).length < 4) return { ok: false, error: "รหัสใหม่ต้องมีอย่างน้อย 4 ตัว" };
  // ห้ามซ้ำกับ Staff PINs
  var staffRows = readStaffList();
  var clash = staffRows.filter(function(s){ return s.active && s.pin === String(d.newPin); })[0];
  if (clash) return { ok: false, error: "PIN นี้ถูกใช้โดยพนักงาน " + clash.name + " อยู่แล้ว" };
  PropertiesService.getScriptProperties().setProperty("ADMIN_PIN_OVERRIDE", String(d.newPin));
  return { ok: true };
}

function handleLogin(phone, pin, deviceId) {
  var adminPin = getAdminPin();
  if (phone === "admin") {
    if (pin === adminPin) return { ok:true, type:"owner", adminName:"เจ้าของ" };
    // Check per-staff PIN from Staff sheet
    var staffRows = readStaffList();
    var matched = staffRows.filter(function(s){ return s.active && s.pin === String(pin); })[0];
    if (matched) return { ok:true, type:"staff", adminName: matched.name, staffId: matched.id };
    if (pin === CONFIG.STAFF_PIN) return { ok:true, type:"staff", adminName:"พนักงาน" };
    return { ok:false, error:"รหัสผ่าน Admin ไม่ถูกต้อง" };
  }
  if (phone === "staff") {
    var staffRows2 = readStaffList();
    var matched2 = staffRows2.filter(function(s){ return s.active && s.pin === String(pin); })[0];
    if (matched2) return { ok:true, type:"staff", adminName: matched2.name, staffId: matched2.id };
    if (pin === CONFIG.STAFF_PIN) return { ok:true, type:"staff", adminName:"พนักงาน" };
    return { ok:false, error:"รหัสผ่าน Staff ไม่ถูกต้อง" };
  }
  var members = readSheet(SHEETS.MEMBERS);
  var idx = -1;
  for (var i=0; i<members.length; i++) {
    var mp = String(members[i].phone || "").replace(/\D/g,"");
    var cp = String(phone || "").replace(/\D/g,"");
    if (mp===cp || "0"+mp===cp || mp==="0"+cp || mp===cp.replace(/^0/,"")) { idx=i; break; }
  }
  if (idx === -1) return { ok:false, error:"ไม่พบเบอร์โทรนี้" };
  var m = members[idx];
  if (String(m.pin) !== String(pin)) return { ok:false, error:"PIN ไม่ถูกต้อง" };
  if (m.deviceId && m.deviceId !== deviceId) return { ok:false, error:"เบอร์นี้ผูกกับอุปกรณ์อื่น แจ้ง Admin" };
  if (!m.deviceId && deviceId) getSheet(SHEETS.MEMBERS).getRange(idx+2,12).setValue(deviceId);
  return { ok:true, type:"user", member:m };
}

// ════════════════════════════════════════════════════════
// 🔒 LOCK HELPER
// ════════════════════════════════════════════════════════
function withLock(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    return fn();
  } catch(e) {
    return { ok:false, error:"ระบบกำลังประมวลผลคำขออื่นอยู่ กรุณาลองใหม่อีกครั้ง" };
  } finally {
    try { lock.releaseLock(); } catch(e2) {}
  }
}

// ════════════════════════════════════════════════════════
// 🕐 BOOKING WINDOW + CANCEL POLICY HELPERS
// ════════════════════════════════════════════════════════
function getSessionDateTime(dateStr, timeStr) {
  var t = formatTimeValue(timeStr);
  var parts = t.split(":");
  var d = new Date(dateStr + "T00:00:00+07:00");
  d.setHours(parseInt(parts[0]), parseInt(parts[1]), 0, 0);
  return d;
}
function checkBookingWindow(sessionDate, classTime) {
  var settings = readSettingsObj();
  var openDays  = parseInt(settings.bookingOpenDaysAhead) || 7;
  var closeHrs  = parseInt(settings.bookingCloseHoursBefore) || 2;
  var sessionDT = getSessionDateTime(sessionDate, classTime);
  var now = new Date();
  var diffMs = sessionDT.getTime() - now.getTime();
  var diffHrs = diffMs / 3600000;
  var diffDays = diffMs / 86400000;
  if (diffHrs < closeHrs) return { ok:false, error:"ปิดรับจองออนไลน์แล้ว (ก่อนเริ่มคลาสน้อยกว่า " + closeHrs + " ชม.) กรุณา Walk-in ที่ฟิตเนส" };
  if (diffDays > openDays) return { ok:false, error:"ยังไม่เปิดรับจองคลาสนี้ (เปิดล่วงหน้า " + openDays + " วัน)" };
  return { ok:true };
}

// ════════════════════════════════════════════════════════
// 🎫 BOOKING
// ════════════════════════════════════════════════════════
function parseList(str) { return str ? String(str).split(",").filter(Boolean) : []; }
// Force text format on a Sheets cell before setValue — prevents Numbers columns from
// treating comma-separated IDs as thousands separators (e.g. "858634748,815319110" → 8.58e+26)
function setTextValue(sh, row, col, val) {
  var r = sh.getRange(row, col);
  r.setNumberFormat("@");
  r.setValue(String(val));
}
function findSessionRow(sid) {
  var sh = getSheet(SHEETS.SESSIONS), data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) { if (data[i][0] === sid) return { row:i+1, data:data[i] }; }
  return null;
}
function countActiveBookings(memberId) {
  var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var sessions = readSheet(SHEETS.SESSIONS);
  var count = 0;
  sessions.forEach(function(s) {
    if (String(s.date) < today) return;
    if (parseList(s.bookings).indexOf(memberId) !== -1) count++;
  });
  return count;
}
function handleBook(d) {
  return withLock(function() {
    var found = findSessionRow(d.sessionId); if (!found) return { ok:false, error:"ไม่พบคลาส" };
    var cls = readSheet(SHEETS.CLASSES).filter(function(c){ return c.id === found.data[1]; })[0];
    if (!cls) return { ok:false, error:"ไม่พบข้อมูลคลาส" };
    if (!d.skipWindowCheck) {
      var win = checkBookingWindow(found.data[2], cls.time);
      if (!win.ok) return win;
    }
    var bk = parseList(found.data[3]);
    if (bk.indexOf(d.memberId) !== -1) return { ok:false, error:"จองแล้ว" };
    if (bk.length >= Number(cls.capacity)) return { ok:false, error:"ที่นั่งเต็ม" };

    var settings = readSettingsObj();
    var maxActive = parseInt(settings.maxActiveBookings) || 0;
    if (maxActive > 0) {
      var activeCount = countActiveBookings(d.memberId);
      if (activeCount >= maxActive) {
        return { ok:false, error:"คุณจองคลาสล่วงหน้าครบ " + maxActive + " คลาสแล้ว ยกเลิกคลาสเก่าก่อนเพื่อจองคลาสใหม่ (กติกานี้ช่วยให้สมาชิกทุกคนมีโอกาสจองเท่าเทียมกัน)" };
      }
    }

    bk.push(d.memberId);
    setTextValue(getSheet(SHEETS.SESSIONS), found.row, 4, bk.join(","));
    appendBooking(d.sessionId, d.memberId, "main");
    deductCredit(d.memberId);

    // ── [v4.3] LINE แจ้งเตือนสมาชิกเมื่อจองสำเร็จ ──────────────────
    try {
      var member = readSheet(SHEETS.MEMBERS).filter(function(m){ return m.id === d.memberId; })[0];
      if (member && member.lineUserId) {
        var classDate = Utilities.formatDate(new Date(found.data[2] + "T00:00:00"), "GMT+7", "d MMMM yyyy");
        var msg = "✅ จองคลาสสำเร็จ!\n\n" +
                  "🏋️ " + cls.name + "\n" +
                  "📅 " + classDate + "\n" +
                  "⏰ " + formatTimeValue(cls.time) + " น.\n" +
                  "👩‍🏫 " + (cls.instructor || "") + "\n" +
                  "🪑 ที่นั่ง " + (bk.length) + "/" + cls.capacity + "\n\n" +
                  "กรุณามาก่อนเวลา 10 นาที 💪\n— Apex Fitness";
        pushLineMessage(member.lineUserId, msg);
      }
    } catch(e) { Logger.log("LINE notify book error: " + e); }

    return { ok:true };
  });
}
function handleJoinWaitlist_v1(d) {
  return withLock(function() {
    var found = findSessionRow(d.sessionId); if (!found) return { ok:false, error:"ไม่พบคลาส" };
    var cls = readSheet(SHEETS.CLASSES).filter(function(c){ return c.id === found.data[1]; })[0];
    if (!cls) return { ok:false, error:"ไม่พบข้อมูลคลาส" };
    var win = checkBookingWindow(found.data[2], cls.time);
    if (!win.ok) return win;
    var wl = parseList(found.data[4]);
    if (wl.indexOf(d.memberId) !== -1) return { ok:false, error:"อยู่ในสำรองแล้ว" };
    if (wl.length >= Number(cls.waitlistCap)) return { ok:false, error:"สำรองเต็ม" };
    wl.push(d.memberId);
    setTextValue(getSheet(SHEETS.SESSIONS), found.row, 5, wl.join(","));
    appendBooking(d.sessionId, d.memberId, "waitlist");

    // ── [v4.3] LINE แจ้งเตือนเมื่ออยู่ในสำรอง ──────────────────────
    try {
      var member = readSheet(SHEETS.MEMBERS).filter(function(m){ return m.id === d.memberId; })[0];
      if (member && member.lineUserId) {
        var classDate2 = Utilities.formatDate(new Date(found.data[2] + "T00:00:00"), "GMT+7", "d MMMM yyyy");
        pushLineMessage(member.lineUserId,
          "🪑 อยู่ในรายชื่อสำรอง\n\n" +
          "🏋️ " + cls.name + "\n" +
          "📅 " + classDate2 + " ⏰ " + formatTimeValue(cls.time) + " น.\n" +
          "ลำดับที่ " + wl.length + "\n\n" +
          "เราจะแจ้งทันทีถ้ามีที่ว่าง 💪\n— Apex Fitness"
        );
      }
    } catch(e) { Logger.log("LINE notify waitlist error: " + e); }

    return { ok:true, position: wl.length };
  });
}
function handleCancel_v1(d) {
  return withLock(function() {
    var found = findSessionRow(d.sessionId); if (!found) return { ok:false, error:"ไม่พบคลาส" };
    var cls = readSheet(SHEETS.CLASSES).filter(function(c){ return c.id === found.data[1]; })[0];
    var settings = readSettingsObj();
    var cancelLimit = parseInt(settings.cancelHoursLimit) || 2;
    var isLateCancel = false;
    if (cls) {
      var sessionDT = getSessionDateTime(found.data[2], cls.time);
      var diffHrs = (sessionDT.getTime() - Date.now()) / 3600000;
      if (diffHrs >= 0 && diffHrs < cancelLimit) isLateCancel = true;
    }
    var bk  = parseList(found.data[3]).filter(function(x){ return x !== d.memberId; });
    var wl  = parseList(found.data[4]);
    var promoted = null;
    if (wl.length > 0 && cls && bk.length < Number(cls.capacity)) {
      promoted = wl.shift(); bk.push(promoted);
      notifyWaitlistPromotion(promoted, d.sessionId, cls ? cls.name : "คลาสนี้");
    }
    var sh = getSheet(SHEETS.SESSIONS);
    setTextValue(sh, found.row, 4, bk.join(","));
    setTextValue(sh, found.row, 5, wl.join(","));

    if (isLateCancel) {
      upsertAttendance(d.sessionId, d.memberId, "noshow");
      var r = incrementNoShow(d.memberId);
      return { ok:true, promoted:promoted, lateCancel:true, noShowResult:r };
    }
    return { ok:true, promoted:promoted, lateCancel:false };
  });
}
function handleCancelWaitlist(d) {
  return withLock(function() {
    var found = findSessionRow(d.sessionId); if (!found) return { ok:false, error:"ไม่พบคลาส" };
    setTextValue(getSheet(SHEETS.SESSIONS), found.row, 5,
      parseList(found.data[4]).filter(function(x){ return x !== d.memberId; }).join(",")
    );
    return { ok:true };
  });
}
function handlePromoteWaitlist(d) {
  return withLock(function() {
    var found = findSessionRow(d.sessionId); if (!found) return { ok:false, error:"ไม่พบคลาส" };
    var bk = parseList(found.data[3]);
    var wl = parseList(found.data[4]).filter(function(x){ return x !== d.memberId; });
    var cls = readSheet(SHEETS.CLASSES).filter(function(c){ return c.id === found.data[1]; })[0];
    if (cls && bk.length >= Number(cls.capacity)) return { ok:false, error:"หลักยังเต็ม" };
    bk.push(d.memberId);
    var sh = getSheet(SHEETS.SESSIONS);
    setTextValue(sh, found.row, 4, bk.join(","));
    setTextValue(sh, found.row, 5, wl.join(","));
    notifyWaitlistPromotion(d.memberId, d.sessionId, cls ? cls.name : "คลาสนี้");
    return { ok:true };
  });
}

// ════════════════════════════════════════════════════════
// 💳 GUEST / WALK-IN BOOKING
// ════════════════════════════════════════════════════════
function handleSubmitGuestBooking(d) {
  var found = findSessionRow(d.sessionId); if (!found) return { ok:false, error:"ไม่พบคลาส" };
  var cls = readSheet(SHEETS.CLASSES).filter(function(c){ return c.id === found.data[1]; })[0];
  if (!cls) return { ok:false, error:"ไม่พบข้อมูลคลาส" };

  var sessionDT = getSessionDateTime(found.data[2], cls.time);
  if (sessionDT.getTime() < Date.now()) return { ok:false, error:"คลาสนี้เริ่มไปแล้ว" };

  var guestId = "G" + Date.now();
  var slipUrl = d.slipBase64 ? saveBase64Image(d.slipBase64, guestId + "_slip.jpg") : "";
  var settings = readSettingsObj();
  var amount = settings.guestPrice || "150";

  return withLock(function() {
    var found2 = findSessionRow(d.sessionId); if (!found2) return { ok:false, error:"ไม่พบคลาส" };
    var bk = parseList(found2.data[3]);
    if (bk.length >= Number(cls.capacity)) return { ok:false, error:"ที่นั่งเต็ม" };

    bk.push(guestId);
    setTextValue(getSheet(SHEETS.SESSIONS), found2.row, 4, bk.join(","));

    var gsh = getSheet(SHEETS.GUEST_BOOKINGS);
    gsh.appendRow([guestId, d.name||"", d.phone||"", d.sessionId, cls.name, found2.data[2], formatTimeValue(cls.time), amount, slipUrl, "pending", new Date().toISOString(), "", ""]);

    return { ok:true, guestId: guestId };
  });
}

function handleConfirmGuestPayment(d) {
  return withLock(function() {
    var sh = getSheet(SHEETS.GUEST_BOOKINGS);
    var data = sh.getDataRange().getValues();
    var rowIdx = -1;
    for (var i=1; i<data.length; i++) { if (data[i][0] === d.guestId) { rowIdx = i; break; } }
    if (rowIdx === -1) return { ok:false, error:"ไม่พบรายการ" };

    var sessionId = data[rowIdx][3];
    var status = d.approve ? "confirmed" : "rejected";
    sh.getRange(rowIdx+1, 10).setValue(status);
    sh.getRange(rowIdx+1, 12).setValue(d.adminName || "");
    sh.getRange(rowIdx+1, 13).setValue(new Date().toISOString());

    if (!d.approve) {
      var found = findSessionRow(sessionId);
      if (found) {
        var bk = parseList(found.data[3]).filter(function(x){ return x !== d.guestId; });
        setTextValue(getSheet(SHEETS.SESSIONS), found.row, 4, bk.join(","));
      }
    }

    // ── [v4.3] LINE แจ้ง Admin เมื่อยืนยัน/ปฏิเสธ payment รายครั้ง ──────────────
    try {
      var guestName  = data[rowIdx][1] || "-";
      var guestPhone = data[rowIdx][2] || "-";
      var clsName    = data[rowIdx][4] || "-";
      var clsDate    = data[rowIdx][5] ? Utilities.formatDate(new Date(data[rowIdx][5] + "T00:00:00"), "GMT+7", "d MMM yyyy") : "-";
      var clsTime    = data[rowIdx][6] || "-";
      var amount     = data[rowIdx][7] || "-";
      var adminMsg = (d.approve
        ? "✅ ยืนยัน payment รายครั้งแล้ว"
        : "❌ ปฏิเสธ payment รายครั้ง") + "\n\n" +
        "👤 " + guestName + " (" + guestPhone + ")\n" +
        "🏋️ " + clsName + "\n" +
        "📅 " + clsDate + " ⏰ " + clsTime + " น.\n" +
        "💵 " + amount + " บาท\n" +
        "👨‍💼 อนุมัติโดย: " + (d.adminName || "-") + "\n\n— Apex Fitness";
      (CONFIG.ADMIN_LINE_USER_IDS || []).forEach(function(uid) {
        try { pushLineMessage(uid, adminMsg); } catch(e2) {}
      });
    } catch(e) { Logger.log("LINE notify confirmGuest error: " + e); }
    // ────────────────────────────────────────────────────────────────────────────────

    return { ok:true };
  });
}

// ════════════════════════════════════════════════════════
// 💬 FEEDBACK / แจ้งปัญหา
// ════════════════════════════════════════════════════════
function handleSubmitFeedback(d) {
  if (!d.message || !String(d.message).trim()) return { ok:false, error:"กรุณากรอกข้อความ" };
  var id = "FB" + Date.now();
  var sh = getSheet(SHEETS.FEEDBACK);
  sh.appendRow([id, d.memberId||"", d.memberName||"", d.type||"feedback", d.message, "open", new Date().toISOString(), "", ""]);
  return { ok:true, id: id };
}
function handleResolveFeedback(d) {
  var sh = getSheet(SHEETS.FEEDBACK);
  var data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) {
    if (data[i][0] === d.feedbackId) {
      sh.getRange(i+1, 6).setValue("resolved");
      sh.getRange(i+1, 8).setValue(d.response || "");
      sh.getRange(i+1, 9).setValue(new Date().toISOString());
      var memberId = data[i][1];
      if (memberId && d.response) {
        notify(memberId, "feedback", "💬 ได้รับการตอบกลับแล้ว", d.response);
        var member = readSheet(SHEETS.MEMBERS).filter(function(m){ return m.id === memberId; })[0];
        if (member && member.lineUserId) pushLineMessage(member.lineUserId, "💬 Apex Fitness ตอบกลับข้อความของคุณ:\n" + d.response);
      }
      return { ok:true };
    }
  }
  return { ok:false, error:"ไม่พบรายการ" };
}

// ════════════════════════════════════════════════════════
function handleCheckin(d) {
  // ── Geofence check (100 m) ──────────────────────────────────
  if (d.lat != null && d.lng != null) {
    var settings = readSettingsObj();
    var gymLat = parseFloat(settings.gymLat || "18.7860278");
    var gymLng = parseFloat(settings.gymLng || "99.0140656");
    var dist = getDistanceMeters(parseFloat(d.lat), parseFloat(d.lng), gymLat, gymLng);
    if (dist > 100) {
      return { ok:false, error:"อยู่ห่างจากฟิตเนส "+Math.round(dist)+" เมตร (ต้องอยู่ภายใน 100 เมตร)" };
    }
  }
  upsertAttendance(d.sessionId, d.memberId, "present");
  return { ok:true };
}
function getDistanceMeters(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var dLat = (lat2-lat1)*Math.PI/180, dLng = (lng2-lng1)*Math.PI/180;
  var a = Math.sin(dLat/2)*Math.sin(dLat/2) +
          Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLng/2)*Math.sin(dLng/2);
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}
function handleMarkAttendance(d) {
  upsertAttendance(d.sessionId, d.memberId, d.status);
  if (d.status === "noshow") { var r = incrementNoShow(d.memberId); return { ok:true, result:r }; }
  return { ok:true };
}
function upsertAttendance(sid, mid, status) {
  var sh = getSheet(SHEETS.ATTENDANCE), data = sh.getDataRange().getValues(), now = new Date().toISOString();
  for (var i=1; i<data.length; i++) { if (data[i][0]===sid && data[i][1]===mid) { sh.getRange(i+1,3).setValue(status); sh.getRange(i+1,4).setValue(now); return; } }
  sh.appendRow([sid, mid, status, now]);
}
function incrementNoShow_v1(mid) {
  var sh = getSheet(SHEETS.MEMBERS), data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) {
    if (data[i][0] === mid) {
      var n = (Number(data[i][8]) || 0) + 1;
      sh.getRange(i+1,9).setValue(n);
      if (n >= CONFIG.NOSHOW_LIMIT && String(data[i][9]).toUpperCase() !== "TRUE") {
        var until = new Date(); until.setDate(until.getDate() + CONFIG.SUSPENSION_DAYS);
        var untilStr = Utilities.formatDate(until, "GMT+7", "yyyy-MM-dd");
        sh.getRange(i+1,10).setValue(true);
        sh.getRange(i+1,11).setValue(untilStr);
        notifySuspension(mid, n, untilStr);
        return { suspended:true, noShowCount:n, until:untilStr };
      }
      notify(mid, "info", "บันทึก No-show", "No-show ครั้งที่ " + n + "/" + CONFIG.NOSHOW_LIMIT);
      return { suspended:false, noShowCount:n };
    }
  }
  return { suspended:false };
}

// ════════════════════════════════════════════════════════
// 👥 MEMBER MANAGEMENT
// ════════════════════════════════════════════════════════
function handleSaveMember(d) {
  var m = d.member, sh = getSheet(SHEETS.MEMBERS), data = sh.getDataRange().getValues();
  var row = [m.id,m.name,m.phone,m.pin||"1234",m.lineId||"",normalizeMembershipType(m.membershipType),m.membershipExpiry||"",m.sessionsLeft||"",m.noShowCount||0,m.suspended||false,m.suspendedUntil||"",m.deviceId||"",m.joinDate||"",m.photoUrl||"",m.lineUserId||""];
  for (var i=1; i<data.length; i++) { if (data[i][0] === m.id) { sh.getRange(i+1,1,1,row.length).setValues([row]); return { ok:true, updated:true }; } }
  sh.appendRow(row);
  return { ok:true, created:true };
}
function handleUploadMemberPhoto(d) {
  if (!d.memberId || !d.photoBase64) return { ok:false, error:"missing data" };
  var url = saveBase64Image(d.photoBase64, d.memberId + "_photo.jpg");
  if (!url) return { ok:false, error:"upload failed" };
  var sh = getSheet(SHEETS.MEMBERS), data = sh.getDataRange().getValues();
  var headers = data[0];
  var col = headers.indexOf("photoUrl");
  if (col === -1) return { ok:false, error:"photoUrl column not found — รัน migrateAddColumns() ก่อน" };
  for (var i=1; i<data.length; i++) {
    if (data[i][0] === d.memberId) { sh.getRange(i+1, col+1).setValue(url); return { ok:true, photoUrl:url }; }
  }
  return { ok:false, error:"ไม่พบสมาชิก" };
}
function handleUnlockDevice(d) {
  var sh = getSheet(SHEETS.MEMBERS), data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) { if (data[i][0] === d.memberId) { sh.getRange(i+1,12).setValue(""); return { ok:true }; } }
  return { ok:false, error:"ไม่พบสมาชิก" };
}

// ✅ [v4.1 FIX] handleSaveSettings
// - ตรวจค่าที่เป็น base64 รูปภาพ → อัปโหลด Drive → เก็บ URL แทน
//   (ป้องกัน Sheets cell overflow: cell limit 50,000 chars, รูป 1400px ≈ 200,000+ chars)
// - return updatedUrls เพื่อให้ Frontend อัปเดต state ด้วย Drive URL
function handleSaveSettings(d) {
  var sh   = getSheet(SHEETS.SETTINGS);
  var data = sh.getDataRange().getValues();
  var updatedUrls = {};

  Object.keys(d.settings || {}).forEach(function(key) {
    var value = d.settings[key];

    // ถ้าค่าเป็น base64 รูปภาพ → อัปโหลด Drive ก่อน เก็บแค่ URL
    if (typeof value === "string" && value.length > 500 && value.indexOf("data:image/") === 0) {
      var url = saveBase64Image(value, key + "_" + Date.now() + ".jpg");
      if (url) {
        updatedUrls[key] = url;
        value = url;
      }
    }

    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(key)) {
        sh.getRange(i + 1, 2).setValue(value);
        data[i][1] = value;
        found = true;
        break;
      }
    }
    if (!found) sh.appendRow([key, value]);
  });

  return { ok: true, updatedUrls: updatedUrls };
}

// ✅ [v4.1 NEW] handleUploadImage
// อัปโหลดรูปภาพเดี่ยวๆ แล้วคืน Drive URL
// Frontend เรียก: postAction("uploadImage", { imageBase64, fileName })
// ใช้สำหรับรูปตารางคลาส / QR รับเงิน
function handleUploadImage(d) {
  if (!d.imageBase64) return { ok: false, error: "missing imageBase64" };
  var fileName = d.fileName || ("img_" + Date.now() + ".jpg");
  var url = saveBase64Image(d.imageBase64, fileName);
  if (!url) return { ok: false, error: "อัปโหลดรูปไม่สำเร็จ ตรวจสอบสิทธิ์ Google Drive" };
  return { ok: true, url: url };
}

// ════════════════════════════════════════════════════════
// 📋 [v4.2] GET SESSION BOOKINGS — ดึงรายชื่อผู้จองในคลาสนั้นๆ
// Frontend เรียก: GET ?action=getSessionBookings&sessionId=c1_2026-06-22
// ════════════════════════════════════════════════════════
function handleGetSessionBookings(sessionId) {
  if (!sessionId) return { ok: false, error: "missing sessionId" };

  var found = findSessionRow(sessionId);
  if (!found) return { ok: false, error: "ไม่พบ session" };

  var memberIds  = parseList(found.data[3]);   // bookings column
  var waitlistIds = parseList(found.data[4]);  // waitlist column

  // สร้าง map id → member object
  var allMembers = readSheet(SHEETS.MEMBERS);
  var memberMap  = {};
  allMembers.forEach(function(m) { memberMap[m.id] = m; });

  // ดึง guest bookings ที่ match sessionId
  var guests = readSheet(SHEETS.GUEST_BOOKINGS).filter(function(g) {
    return String(g.sessionId) === String(sessionId);
  }).map(function(g) {
    return {
      id:   g.id,
      name: g.name || "(Walk-in)",
      phone: g.phone || "",
      type: "guest",
      status: g.status,   // pending / confirmed / rejected
      slipUrl: g.slipUrl || "",
    };
  });

  var booked = memberIds.map(function(mid) {
    // Guest id เริ่มด้วย "G" ไม่ต้องหาใน Members sheet
    if (mid.charAt(0) === "G") return null;
    var m = memberMap[mid];
    if (!m) return { id: mid, name: "ไม่พบสมาชิก", phone: "", type: "member" };
    return {
      id:   m.id,
      name: m.name,
      phone: m.phone,
      membershipType: normalizeMembershipType(m.membershipType),
      type: "member",
    };
  }).filter(Boolean);

  var waitlist = waitlistIds.map(function(mid) {
    var m = memberMap[mid];
    if (!m) return { id: mid, name: "ไม่พบสมาชิก", phone: "", type: "member" };
    return {
      id:   m.id,
      name: m.name,
      phone: m.phone,
      membershipType: normalizeMembershipType(m.membershipType),
      type: "member",
    };
  });

  return {
    ok:       true,
    sessionId: sessionId,
    booked:   booked,
    waitlist: waitlist,
    guests:   guests,
    total:    booked.length + guests.filter(function(g){ return g.status !== "rejected"; }).length,
  };
}

// ════════════════════════════════════════════════════════
// 📤 [v4.2] EXPORT REPORT — สร้าง Google Spreadsheet ใหม่ ส่ง URL กลับ
// Frontend เรียก: postAction("exportReport", { type: "members"|"bookings"|"attendance"|"guestBookings" })
// ════════════════════════════════════════════════════════
function handleExportReport(d) {
  var type  = d.type || "members";
  var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var title = "Apex Fitness — " + type + " — " + today;

  var ss = SpreadsheetApp.create(title);
  var sh = ss.getActiveSheet();
  sh.setName(type);

  var headers, rows;

  if (type === "members") {
    headers = ["ID","ชื่อ","เบอร์โทร","ประเภทสมาชิก","วันหมดอายุ","ครั้งคงเหลือ","No-show","ระงับ","ระงับถึง","วันเข้าร่วม"];
    rows = readSheet(SHEETS.MEMBERS).map(function(m) {
      return [m.id, m.name, m.phone, normalizeMembershipType(m.membershipType),
              m.membershipExpiry||"", m.sessionsLeft||"", m.noShowCount||0,
              m.suspended||false, m.suspendedUntil||"", m.joinDate||""];
    });

  } else if (type === "bookings") {
    headers = ["ID จอง","Session","สมาชิก ID","ประเภท","วันที่จอง"];
    rows = readSheet(SHEETS.BOOKINGS).map(function(b) {
      return [b.id, b.sessionId, b.memberId, b.type, b.createdAt];
    });

  } else if (type === "attendance") {
    // join ชื่อสมาชิกเข้าไปด้วย
    var memberMap2 = {};
    readSheet(SHEETS.MEMBERS).forEach(function(m) { memberMap2[m.id] = m.name; });
    headers = ["Session","สมาชิก ID","ชื่อ","สถานะ","เวลาบันทึก"];
    rows = readSheet(SHEETS.ATTENDANCE).map(function(a) {
      return [a.sessionId, a.memberId, memberMap2[a.memberId]||"", a.status, a.markedAt];
    });

  } else if (type === "guestBookings") {
    headers = ["ID","ชื่อ","เบอร์โทร","Session","คลาส","วันที่","เวลา","ยอด (บาท)","สลิป","สถานะ","วันที่สร้าง","ยืนยันโดย"];
    rows = readSheet(SHEETS.GUEST_BOOKINGS).map(function(g) {
      return [g.id, g.name, g.phone, g.sessionId, g.className, g.classDate,
              g.classTime, g.amount, g.slipUrl, g.status, g.createdAt, g.confirmedBy||""];
    });

  } else {
    return { ok: false, error: "type ไม่รองรับ: " + type };
  }

  // เขียน header
  if (headers.length > 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#F97316").setFontColor("#FFFFFF");
  }
  // เขียน data
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sh.autoResizeColumns(1, headers.length);
  }
  sh.setFrozenRows(1);

  // แชร์ link (view only) เพื่อให้แอดมินเปิดได้
  var file = DriveApp.getFileById(ss.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log("📤 Export '" + type + "' → " + ss.getUrl());
  return { ok: true, url: ss.getUrl(), rows: rows.length, title: title };
}

// ════════════════════════════════════════════════════════
// 🔄 SYNC จาก Aristo
// ════════════════════════════════════════════════════════
function normalizeMembershipType(t) {
  var s = String(t || "").trim().toLowerCase();
  if (s === "monthly" || s.indexOf("month") !== -1 || s.indexOf("เดือน") !== -1) return "monthly";
  if (s === "per_session" || s === "persession" || s.indexOf("session") !== -1 || s.indexOf("ครั้ง") !== -1) return "per_session";
  return s || "monthly";
}
function handleSyncAristo(d) {
  var members = d.members || [];
  if (!members.length) return { ok:false, error:"no members" };
  var sh = getSheet(SHEETS.MEMBERS);
  var existing = readSheet(SHEETS.MEMBERS);
  var saved = {};
  existing.forEach(function(e) {
    var key = String(e.phone || "").replace(/\D/g,"");
    if (key) saved[key] = { pin: e.pin||"1234", deviceId: e.deviceId||"", noShow: e.noShowCount||"0", photoUrl: e.photoUrl||"", lineUserId: e.lineUserId||"" };
  });
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2,1,lastRow-1,15).clearContent();
  var rows = [];
  members.forEach(function(m) {
    if (!m.phone || String(m.phone).replace(/\D/g,"").length < 9) return;
    var key = String(m.phone).replace(/\D/g,"");
    var old = saved[key] || {};
    rows.push([m.id, m.name, m.phone, old.pin||"1234", m.lineId||"",
      normalizeMembershipType(m.membershipType), m.membershipExpiry||"", m.sessionsLeft||"",
      old.noShow||"0", m.suspended||"FALSE", m.suspendedUntil||"", old.deviceId||"", m.joinDate||"",
      old.photoUrl||"", old.lineUserId||""]);
  });
  if (rows.length > 0) sh.getRange(2,1,rows.length,15).setValues(rows);
  Logger.log("🔄 Sync Aristo เสร็จ: " + rows.length + " คน");
  return { ok:true, total:rows.length };
}

// ════════════════════════════════════════════════════════
// ⚙️ UTILITIES
// ════════════════════════════════════════════════════════
function appendBooking(sid, mid, type) {
  getSheet(SHEETS.BOOKINGS).appendRow(["BK"+Date.now(), sid, mid, type, new Date().toISOString()]);
}
function deductCredit(mid) {
  var sh = getSheet(SHEETS.MEMBERS), data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) {
    if (data[i][0] === mid && data[i][5] === "per_session") {
      sh.getRange(i+1,8).setValue(Math.max(0,(Number(data[i][7])||0)-1)); return;
    }
  }
}
function notify(mid, type, title, body) {
  getSheet(SHEETS.NOTIFICATIONS).appendRow(["N"+Date.now()+Math.floor(Math.random()*1000), mid, type, title, body, new Date().toISOString(), false]);
}

// ════════════════════════════════════════════════════════
// ⏰ ตั้ง Time Triggers
// ════════════════════════════════════════════════════════
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("generateSessionsForNext14Days").timeBased().everyDays(1).atHour(0).create();
  ScriptApp.newTrigger("sendDailyAdminSummary").timeBased().everyDays(1).atHour(20).create();
  ScriptApp.newTrigger("sendExpiryReminders").timeBased().everyDays(1).atHour(9).create();
  Logger.log("✅ Triggers setup done");
  try { SpreadsheetApp.getUi().alert("✅ ตั้ง Triggers สำเร็จ!"); } catch(e) {}
}
