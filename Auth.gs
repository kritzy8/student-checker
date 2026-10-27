/******** Auth.gs — ยืนยันตัวตนผ่าน Bridge (ใช้กับ @gmail.com ได้) ********/

// ใส่ URL /exec ของโปรเจกต์ Bridge (ได้มาหลังทำส่วน B)
const BRIDGE_URL =
  "https://script.google.com/macros/s/AKfycbzySE16f8Z_ixAxwzwbpap_wU5UbBXizqgqubyJzjlf3S_ydKRFC5zl-Gxtj_UCVgmtAA/exec";

let CURRENT_EMAIL_ = "";

function currentEmail_() {
  return CURRENT_EMAIL_;
}

/** รันครั้งเดียวจาก editor: สร้างรหัสลับ แล้วคัดลอกค่าใน Log ไปใส่ใน Bridge */
function createAuthSecret() {
  const s =
    Utilities.getUuid().replace(/-/g, "") +
    Utilities.getUuid().replace(/-/g, "");
  PropertiesService.getScriptProperties().setProperty("AUTH_SECRET", s);
  Logger.log("AUTH_SECRET = " + s);
}

function authKey_() {
  const k = PropertiesService.getScriptProperties().getProperty("AUTH_SECRET");
  if (!k)
    throw new Error(
      "ยังไม่ได้ตั้งค่า AUTH_SECRET (ให้รันฟังก์ชัน createAuthSecret ก่อน)",
    );
  return k;
}

function authSign_(s) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(s, authKey_()),
  );
}

/** ตรวจตั๋วจาก Bridge: ถูกต้องคืนอีเมล, ไม่ถูกต้อง/หมดอายุคืน '' */
function verifyAuth_(auth) {
  const parts = String(auth || "").split(".");
  if (parts.length !== 2) return "";
  let payload;
  try {
    payload = Utilities.newBlob(
      Utilities.base64DecodeWebSafe(parts[0]),
    ).getDataAsString();
  } catch (err) {
    return "";
  }
  if (authSign_(payload) !== parts[1]) return "";
  const i = payload.lastIndexOf("|");
  const email = payload.substring(0, i);
  const exp = Number(payload.substring(i + 1));
  if (!email || !exp || Date.now() > exp) return "";
  return email;
}

function buildQuery_(obj) {
  return Object.keys(obj)
    .filter(function (k) {
      return obj[k] !== undefined && obj[k] !== "";
    })
    .map(function (k) {
      return encodeURIComponent(k) + "=" + encodeURIComponent(obj[k]);
    })
    .join("&");
}

/** หน้าเปลี่ยนเส้นทาง (มีปุ่มสำรองเผื่อเบราว์เซอร์บล็อกการเปลี่ยนหน้าอัตโนมัติ) */
function goTopPage_(url, message, buttonText) {
  const href = String(url).replace(/&/g, "&amp;").replace(/"/g, "&quot;");
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;margin-top:40px">' +
      "<p>" +
      message +
      "</p>" +
      '<a href="' +
      href +
      '" target="_top" style="display:inline-block;font-size:20px;padding:12px 24px;' +
      'background:#2563eb;color:#fff;border-radius:10px;text-decoration:none">' +
      buttonText +
      "</a></div>" +
      "<script>try{window.top.location.href=" +
      JSON.stringify(url) +
      "}catch(e){}</script>",
  ).addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function redirectToBridge_(params) {
  const q = Object.assign({}, params);
  delete q.auth;
  const qs = buildQuery_(q);
  return goTopPage_(
    BRIDGE_URL + (qs ? "?" + qs : ""),
    "กำลังยืนยันตัวตน…",
    "กดที่นี่เพื่อเข้าสู่ระบบ",
  );
}

/** ทางเข้าเดียวสำหรับหน้าเว็บทุกหน้า: ตรวจตั๋วก่อน แล้วค่อยเรียกฟังก์ชันจริง */
function api(auth, fn, args) {
  const email = verifyAuth_(auth);
  if (!email) throw new Error("เซสชันหมดอายุ กรุณาเปิดลิงก์ใหม่อีกครั้ง");
  const FN = {
    getInit: getInit,
    startSession: startSession,
    getToken: getToken,
    getStatus: getStatus,
    stopSession: stopSession,
    markManual: markManual,
    removeAttendance: removeAttendance,
    checkin: checkin,
    getAdminInit: getAdminInit,
    getReport: getReport,
    setStatus: setStatus,
    voidSession: voidSession,
    getDashboard: getDashboard,
    //จัดการตารางสอน
    saveSchedule: saveSchedule,
    deleteSchedule: deleteSchedule,
    //จัดการรายชื่อนักเรียน
    getStudentsPage: getStudentsPage,
    addStudent: addStudent,
    updateStudent: updateStudent,
    deleteStudent: deleteStudent,
    //exportเป็นexcel
    downloadExcelReport: downloadExcelReport,
  };
  if (!Object.prototype.hasOwnProperty.call(FN, fn))
    throw new Error("ไม่อนุญาตให้เรียกฟังก์ชันนี้");
  CURRENT_EMAIL_ = email;
  return FN[fn].apply(null, Array.isArray(args) ? args : []);
}
