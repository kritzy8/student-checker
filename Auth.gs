/******** Auth.gs — ยืนยันตัวตนผ่าน Bridge (ใช้กับ @gmail.com ได้) ********/

// ใส่ URL /exec ของโปรเจกต์ Bridge (ได้มาหลังทำส่วน B)
const BRIDGE_URL =
  "https://script.google.com/macros/s/AKfycbzySE16f8Z_ixAxwzwbpap_wU5UbBXizqgqubyJzjlf3S_ydKRFC5zl-Gxtj_UCVgmtAA/exec";

let CURRENT_EMAIL_ = "";

function currentEmail_() {
  return CURRENT_EMAIL_;
}

/** รันครั้งเดียวจาก editor: สร้างรหัสลับ แล้วคัดลอกค่าใน Log ไปใส่ใน Bridge (จะขึ้นเตือนหากมีค่าอยู่แล้ว)*/
function createAuthSecret() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("AUTH_SECRET")) {
    throw new Error(
      "มี AUTH_SECRET อยู่แล้ว จะไม่สร้างใหม่ทับ (ถ้าต้องการเปลี่ยนจริง ๆ ให้ลบค่าเดิมใน Script Properties ก่อน)",
    );
  }
  const s =
    Utilities.getUuid().replace(/-/g, "") +
    Utilities.getUuid().replace(/-/g, "");
  props.setProperty("AUTH_SECRET", s);
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
    '<!doctype html><html lang="th"><head><base target="_top">' +
      '<meta charset="utf-8">' +
      "<style>" +
      'body{margin:0;font-family:"Noto Sans Thai",system-ui,sans-serif;background:#f4f6fb;color:#14213d;' +
      "display:flex;align-items:center;justify-content:center;min-height:90vh;cursor:pointer;-webkit-tap-highlight-color:transparent}" +
      ".box{background:#fff;border:1px solid #e3e8f2;border-radius:20px;padding:28px 24px;max-width:360px;width:90%;text-align:center;box-shadow:0 10px 28px rgba(20,33,61,.08)}" +
      ".spin{width:28px;height:28px;border:3px solid #e8effd;border-top-color:#2563eb;border-radius:50%;margin:0 auto 14px;animation:sp .7s linear infinite}" +
      "@keyframes sp{to{transform:rotate(360deg)}}" +
      ".title{font-size:16px;font-weight:700;margin:0 0 6px;color:#14213d}" +
      ".msg{font-size:13.5px;color:#64708a;margin:0 0 12px;word-break:break-all}" +
      ".sub{font-size:13px;color:#64708a;margin:0 0 18px;line-height:1.5}" +
      ".btn{display:block;width:100%;box-sizing:border-box;font-size:17px;font-weight:700;padding:13px 20px;" +
      "background:#2563eb;color:#fff;border-radius:8px;text-decoration:none}" +
      "</style></head><body>" +
      '<div class="box">' +
      '<div class="spin"></div>' +
      '<p class="title">กำลังพาไปยังหน้าเว็บ...</p>' +
      (message ? '<p class="msg">' + message + "</p>" : "") +
      '<p class="sub">หากหน้าจอไม่เปลี่ยนโดยอัตโนมัติ กรุณากดปุ่มด้านล่าง</p>' +
      '<a id="goLink" class="btn" href="' +
      href +
      '" target="_top">' +
      buttonText +
      " ›</a>" +
      "</div>" +
      "<script>" +
      "var u=" +
      JSON.stringify(url) +
      ";" +
      "function go(){" +
      '  try{var a=document.getElementById("goLink");if(a)a.click();}catch(e){}' +
      "  try{window.top.location.replace(u);}catch(e){}" +
      '  try{window.open(u,"_top");}catch(e){}' +
      "}" +
      "go();" +
      "setTimeout(go,150);" +
      'document.addEventListener("touchstart",function(){var a=document.getElementById("goLink");if(a)a.click();},{once:true});' +
      "</script></body></html>",
  )
    .setTitle("กำลังพาไปยังหน้าเว็บ...")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
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
    markAllPresent: markAllPresent,
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
    //หมายเหตุนักเรียน
    getStudentNotes: getStudentNotes,
    saveStudentNote: saveStudentNote,
  };
  if (!Object.prototype.hasOwnProperty.call(FN, fn))
    throw new Error("ไม่อนุญาตให้เรียกฟังก์ชันนี้");
  CURRENT_EMAIL_ = email;
  return FN[fn].apply(null, Array.isArray(args) ? args : []);
}
