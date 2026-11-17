/******** Bridge — อ่านอีเมลผู้ใช้ แล้วออกตั๋วส่งต่อไปแอปหลัก ********/

// ใส่ URL /exec ของแอปหลัก
const MAIN_URL =
  "https://script.google.com/macros/s/AKfycbyftTVG7lUWAHO6SFUOpqYDKoM_HBTPrrmGxpRsKo0lAcvUO3X1rT6We1ZthJt7mhblTg/exec";

function doGet(e) {
  const p = (e && e.parameter) || {};
  const email = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!email) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px">อ่านอีเมลไม่ได้ กรุณาล็อกอินบัญชี Google แล้วลองใหม่</p>',
    );
  }
  // หน้าเช็กชื่อ: ตั๋วมีอายุ 10 นาที / หน้าอื่น: 8 ชั่วโมง
  const ttl = p.p === "c" ? 10 * 60 * 1000 : 8 * 60 * 60 * 1000;
  const payload = email + "|" + (Date.now() + ttl);
  const auth =
    Utilities.base64EncodeWebSafe(payload) + "." + authSign_(payload);

  const q = Object.assign({}, p, { auth: auth });
  return goTopPage_(
    MAIN_URL + "?" + buildQuery_(q),
    "ยืนยันตัวตนแล้ว: " + escHtml_(email),
    "ดำเนินการต่อ",
  );
}

function authSign_(s) {
  const k = PropertiesService.getScriptProperties().getProperty("AUTH_SECRET");
  if (!k) throw new Error("Bridge ยังไม่ได้ตั้งค่า AUTH_SECRET");
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(s, k),
  );
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

function escHtml_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    }[c];
  });
}

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

function checkSecret() {
  const k = PropertiesService.getScriptProperties().getProperty("AUTH_SECRET");
  Logger.log(
    k
      ? k.length +
          " ตัวอักษร, ขึ้นต้น " +
          k.slice(0, 6) +
          ", ลงท้าย " +
          k.slice(-6)
      : "ไม่มีค่า (MISSING)",
  );
}
