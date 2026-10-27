/******** Bridge — อ่านอีเมลผู้ใช้ แล้วออกตั๋วส่งต่อไปแอปหลัก ********/

// ใส่ URL /exec ของแอปหลัก
const MAIN_URL =
  "https://script.google.com/macros/s/AKfycbzuIgaG0izNrNIzpAEmhiK7GFX_65f8AvYELq1lc9J0dnGdfUERFtMBjj0mTTjpU5_YKg/exec";

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
