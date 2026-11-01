/**
 * ระบบเช็กชื่อด้วย QR หมุนเวียน — Google Apps Script (ต้นแบบ)
 * ไฟล์ในโปรเจกต์: Code.gs, Auth.gs, Home.html, Teacher.html, Student.html, Admin.html
 *
 * หลักการ:
 *  - อาจารย์เปิดหน้า Teacher บนโปรเจกเตอร์ QR เปลี่ยนทุก SLOT_SECONDS วินาที
 *  - ใน QR มีโทเค็น HMAC ที่ผูกกับ session + ช่วงเวลา ภาพที่ถ่ายส่งต่อจะใช้ไม่ได้เมื่อหมดอายุ
 *  - นักเรียนต้องล็อกอินอีเมลโรงเรียน ระบบเทียบกับชีต Students เอง (ไม่ต้องพิมพ์ชื่อ)
 *  - 1 บัญชีเช็กได้ครั้งเดียวต่อ session และตรวจจับอุปกรณ์เดียวกันที่เช็กหลายบัญชี
 */

const CONFIG = {
  SLOT_SECONDS: 15, // QR เปลี่ยนทุกกี่วินาที
  GRACE_SLOTS: 2, // ยอมรับโทเค็นของรอบก่อนหน้ากี่รอบ (กันสแกนช้าและเผื่อเวลาระบบทำงาน) => อายุจริง 15-30 วินาที
  DEFAULT_MINUTES: 5, // เวลาเปิดรับเช็กชื่อ
  MAX_MINUTES: 30,
  TZ: "Asia/Bangkok",
  CACHE_STUDENTS_SEC: 600,
  CACHE_SESSION_SEC: 21600,
};

const SHEETS = {
  STUDENTS: "Students",
  TEACHERS: "Teachers",
  SESSIONS: "Sessions",
  ATTENDANCE: "Attendance",
  ADMINS: "Admins",
  ADJUSTMENTS: "Adjustments",
  SCHEDULE: "Schedule",
};

const HEADERS = {
  Students: ["email", "studentId", "name", "class"],
  Teachers: ["email", "name", "fullName"],
  Sessions: [
    "sessionId",
    "class",
    "teacher",
    "start",
    "end",
    "secret",
    "status",
  ],
  Attendance: [
    "timestamp",
    "sessionId",
    "class",
    "email",
    "studentId",
    "name",
    "deviceId",
    "method",
    "flag",
  ],
  Admins: ["email", "name", "fullName"],
  Adjustments: ["sessionId", "email", "status", "note", "by", "updatedAt"],
  Schedule: [
    "entryId",
    "teacher",
    "day",
    "period",
    "class",
    "subject",
    "room",
    "updatedAt",
  ],
};

// รหัสสถานะในรายงาน: p=มาเรียน l=มาสาย v=ลา a=ขาด
const STATUS_CODE = { present: "p", late: "l", leave: "v", absent: "a" };

/* ===================== ติดตั้งครั้งแรก ===================== */

/** รันครั้งเดียวจากตัวแก้ไข Apps Script (ผูกกับ Google Sheets ที่เปิดอยู่) */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty("SS_ID", ss.getId());
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const head = HEADERS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
      sh.setFrozenRows(1);
    } else {
      // ชีตเดิม: เติมหัวคอลัมน์ที่ยังขาด (เช่น fullName) โดยไม่แตะข้อมูล
      const cur = sh.getRange(1, 1, 1, head.length).getValues()[0];
      head.forEach(function (h, i) {
        if (!cur[i])
          sh.getRange(1, i + 1)
            .setValue(h)
            .setFontWeight("bold");
      });
    }
  });
  const me = (Session.getActiveUser().getEmail() || "").toLowerCase();
  const teachers = ss.getSheetByName(SHEETS.TEACHERS);
  if (me && teachers.getLastRow() < 2) teachers.appendRow([me, "ผู้ดูแลระบบ"]);
  const admins = ss.getSheetByName(SHEETS.ADMINS);
  if (me && admins.getLastRow() < 2) admins.appendRow([me, "ผู้ดูแลระบบ"]);
  ss.getSheetByName(SHEETS.SESSIONS).hideColumns(6);
  // ตารางสอน: class/subject/room เป็นข้อความล้วน
  ss.getSheetByName(SHEETS.SCHEDULE).getRange("E:G").setNumberFormat("@");
  // รายชื่อนักเรียน: ทุกคอลัมน์เป็นข้อความล้วน (เก็บเลขศูนย์นำหน้า เช่น 0123 ได้)
  ss.getSheetByName(SHEETS.STUDENTS).getRange("A:D").setNumberFormat("@");
  clearCache();
}

/** (ไม่บังคับ) สร้างนักเรียนจำลอง 5 ห้อง x 40 คน สำหรับทดสอบหน้าจอและโหลด */
function seedDemo() {
  const sh = sheet_(SHEETS.STUDENTS);
  if (sh.getLastRow() > 1) throw new Error("ชีต Students มีข้อมูลอยู่แล้ว");
  const rows = [];
  for (let c = 1; c <= 5; c++) {
    for (let n = 1; n <= 40; n++) {
      const nn = ("0" + n).slice(-2);
      rows.push([
        "demo-" + c + "-" + nn + "@example.com",
        "4" + c + nn,
        "นักเรียนตัวอย่าง " + c + "/" + nn,
        "ม.4/" + c,
      ]);
    }
  }
  sh.getRange(2, 1, rows.length, 4).setValues(rows);
  clearCache();
}

/* ===================== ข้อมูลทดสอบ "ประวัติการเช็กชื่อ" (แยกจาก seedDemo) =====================
 * seedDemoHistory()   สร้างนักเรียนจำลอง + รอบเช็กชื่อย้อนหลัง + บันทึกการเข้าเรียน + รายการที่แอดมินปรับแก้
 * removeDemoHistory() ลบเฉพาะข้อมูลที่สร้างโดยฟังก์ชันด้านบน (ดูจากอีเมล demohist-* และรหัสรอบ dh*) ข้อมูลจริงไม่ถูกแตะ
 * ปรับค่าได้ที่ DEMO_HIST ด้านล่าง แล้วรันจากตัวแก้ไข Apps Script
 */
const DEMO_HIST = {
  CLASS: "ม.4/2", // ห้องที่จะสร้างข้อมูลทดสอบ
  STUDENTS: 40, // จำนวนนักเรียนจำลอง
  DAYS_BACK: 56, // ย้อนหลังกี่วัน (นับเฉพาะจันทร์-ศุกร์ ไม่รวมวันนี้)
  PERIODS: [1, 3, 5], // คาบที่เปิดเช็กชื่อในแต่ละวัน (เวลาเริ่มตามตารางคาบ)
  OPEN_MINUTES: 5, // ระยะเวลาเปิดรับเช็กชื่อ
  EMAIL_PREFIX: "demohist-",
  SESSION_PREFIX: "dh",
};

function demoRand_(seed) {
  let x = seed >>> 0;
  return function () {
    x = (Math.imul(x, 1664525) + 1013904223) >>> 0;
    return x / 4294967296;
  };
}

function seedDemoHistory() {
  const D = DEMO_HIST;
  const emailP = D.EMAIL_PREFIX;
  if (
    rows_(SHEETS.STUDENTS).some(function (r) {
      return String(r.email).indexOf(emailP) === 0;
    })
  ) {
    throw new Error("มีข้อมูลทดสอบอยู่แล้ว กรุณารัน removeDemoHistory() ก่อน");
  }
  const rnd = demoRand_(2569);
  const teacher =
    (Session.getActiveUser().getEmail() || "").toLowerCase() ||
    String((rows_(SHEETS.ADMINS)[0] || {}).email || "").toLowerCase() ||
    "demo@example.com";

  // 1) นักเรียนจำลอง (แต่ละคนมีนิสัยการมาเรียนต่างกัน)
  const firstNames = [
    "สมชาย",
    "สมหญิง",
    "วิชัย",
    "มาลี",
    "ประเสริฐ",
    "สุดา",
    "ธนากร",
    "พิมพ์ใจ",
    "กิตติ",
    "นภา",
  ];
  const lastNames = [
    "ใจดี",
    "รักเรียน",
    "สุขสันต์",
    "แสนดี",
    "มั่นคง",
    "เจริญผล",
    "บุญมา",
    "ศรีสุข",
  ];
  const stu = [];
  for (let n = 1; n <= D.STUDENTS; n++) {
    const nn = ("0" + n).slice(-2);
    stu.push({
      email: emailP + nn + "@example.com",
      studentId: "D" + nn,
      name:
        firstNames[(n - 1) % firstNames.length] +
        " " +
        lastNames[Math.floor((n - 1) / firstNames.length) % lastNames.length] +
        " (ทดสอบ " +
        nn +
        ")",
      rate: n % 9 === 0 ? 0.45 + rnd() * 0.2 : 0.8 + rnd() * 0.2, // บางคนมาน้อย
    });
  }
  const shS = sheet_(SHEETS.STUDENTS);
  const startRowS = shS.getLastRow() + 1;
  const rangeS = shS.getRange(startRowS, 1, stu.length, 4);
  rangeS.setNumberFormat("@");
  rangeS.setValues(
    stu.map(function (x) {
      return [x.email, x.studentId, x.name, D.CLASS];
    }),
  );

  // 2) วันเรียน (จันทร์-ศุกร์ ย้อนหลัง ไม่รวมวันนี้) x คาบ
  const times = {};
  PERIOD_TIMES.forEach(function (p) {
    times[p[0]] = p[1];
  });
  const sessions = [];
  const att = [];
  const adj = [];
  const notesPool = [
    "ลาป่วย มีใบรับรองแพทย์",
    "ลากิจ",
    "ไปแข่งขันกีฬา",
    "มาสายเพราะรถติด",
    "ติดภารกิจโรงเรียน",
  ];
  const today = new Date();
  let counter = 0;
  for (let back = D.DAYS_BACK; back >= 1; back--) {
    const day = new Date(today.getTime() - back * 86400000);
    const ymd = Utilities.formatDate(day, CONFIG.TZ, "yyyy-MM-dd");
    const dow = Number(Utilities.formatDate(day, CONFIG.TZ, "u"));
    if (dow > 5) continue;
    D.PERIODS.forEach(function (pn) {
      const startMs = new Date(ymd + "T" + times[pn] + ":00+07:00").getTime();
      const sid = D.SESSION_PREFIX + ("000" + ++counter).slice(-4);
      sessions.push([
        sid,
        D.CLASS,
        teacher,
        new Date(startMs),
        new Date(startMs + D.OPEN_MINUTES * 60000),
        Utilities.getUuid(),
        "closed",
      ]);
      stu.forEach(function (x) {
        if (rnd() < x.rate) {
          const t = startMs + Math.floor(10000 + rnd() * 240000);
          att.push([
            new Date(t),
            sid,
            D.CLASS,
            x.email,
            x.studentId,
            x.name,
            "demo-dev-" + x.studentId,
            "qr",
            "",
          ]);
          if (rnd() < 0.04)
            adj.push([
              sid,
              x.email,
              "late",
              "มาสายเพราะรถติด",
              teacher,
              new Date(startMs),
            ]);
        } else {
          const r = rnd();
          if (r < 0.25)
            adj.push([
              sid,
              x.email,
              "leave",
              notesPool[Math.floor(rnd() * 3)],
              teacher,
              new Date(startMs),
            ]);
          else if (r < 0.3)
            adj.push([
              sid,
              x.email,
              "present",
              "ติดภารกิจโรงเรียน",
              teacher,
              new Date(startMs),
            ]);
        }
      });
    });
  }

  const put = function (name, rows, textCols) {
    if (!rows.length) return;
    const sh = sheet_(name);
    const rg = sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length);
    if (textCols) rg.setNumberFormat("@");
    rg.setValues(rows);
  };
  put(SHEETS.SESSIONS, sessions);
  put(SHEETS.ATTENDANCE, att);
  put(SHEETS.ADJUSTMENTS, adj);
  clearCache();
  return {
    class: D.CLASS,
    students: stu.length,
    sessions: sessions.length,
    attendance: att.length,
    adjustments: adj.length,
  };
}

// ลบแถวที่คอลัมน์ col (เริ่มที่ 1) ขึ้นต้นด้วย prefix โดยลบเป็นช่วงต่อเนื่องจากล่างขึ้นบน
function deleteRowsByPrefix_(name, col, prefix) {
  const sh = sheet_(name);
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const vals = sh.getRange(2, col, n - 1, 1).getValues();
  const hit = [];
  for (let i = 0; i < vals.length; i++)
    if (String(vals[i][0]).indexOf(prefix) === 0) hit.push(i + 2);
  let removed = 0;
  for (let k = hit.length - 1; k >= 0; ) {
    let j = k;
    while (j > 0 && hit[j - 1] === hit[j] - 1) j--;
    sh.deleteRows(hit[j], k - j + 1);
    removed += k - j + 1;
    k = j - 1;
  }
  return removed;
}

function removeDemoHistory() {
  const D = DEMO_HIST;
  const out = {
    adjustments: deleteRowsByPrefix_(SHEETS.ADJUSTMENTS, 1, D.SESSION_PREFIX),
    attendance: deleteRowsByPrefix_(SHEETS.ATTENDANCE, 2, D.SESSION_PREFIX),
    sessions: deleteRowsByPrefix_(SHEETS.SESSIONS, 1, D.SESSION_PREFIX),
    students: deleteRowsByPrefix_(SHEETS.STUDENTS, 1, D.EMAIL_PREFIX),
  };
  clearCache();
  return out;
}

/** รันเมื่อแก้ไขชีต Students/Teachers แล้วอยากให้มีผลทันที (ปกติหมดอายุเองใน 5-10 นาที) */
function clearCache() {
  CacheService.getScriptCache().removeAll(["students", "teachers", "admins"]);
}

/* ===================== หน้าเว็บ ===================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const email = verifyAuth_(p.auth);
  if (!email) return redirectToBridge_(p);
  CURRENT_EMAIL_ = email;

  // หน้าให้นักเรียนเปิดครั้งแรก เพื่อกดอนุญาตบัญชีไว้ล่วงหน้า
  if (p.p === "ok") {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px;text-align:center;font-size:18px">' +
        "✓ เชื่อมบัญชี " +
        email +
        " เรียบร้อย<br>พร้อมสแกน QR เช็กชื่อแล้ว</p>",
    )
      .setTitle("พร้อมใช้งาน")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (p.p === "admin") {
    if (!isAdmin_(email)) {
      return HtmlService.createHtmlOutput(
        '<p style="font-family:sans-serif;padding:24px">ไม่มีสิทธิ์เข้าถึงหน้านี้ (เฉพาะผู้ดูแลระบบ)</p>',
      ).setTitle("ไม่มีสิทธิ์");
    }
    const ta = HtmlService.createTemplateFromFile("Admin");
    ta.auth = p.auth;
    return ta
      .evaluate()
      .setTitle("ระบบหลังบ้าน")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (p.p === "c") {
    const safe = function (v) {
      return /^[A-Za-z0-9._-]{0,64}$/.test(v || "") ? v || "" : "";
    };
    const t = HtmlService.createTemplateFromFile("Student");
    t.sessionId = safe(p.s);
    t.token = safe(p.t);
    t.auth = p.auth;
    return t
      .evaluate()
      .setTitle("เช็กชื่อ")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // หน้าเช็กชื่อของอาจารย์ (Teacher)
  if (p.p === "t") {
    const tt = HtmlService.createTemplateFromFile("Teacher");
    tt.auth = p.auth;
    tt.bridgeUrl = BRIDGE_URL; // ใหม่
    tt.canAdmin = isAdmin_(email); // ใหม่
    return tt
      .evaluate()
      .setTitle("เช็กชื่อด้วย QR")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // หน้าแรก: เมนูให้เลือกไปหน้าเช็กชื่อ หรือหน้าสรุปผลของผู้ดูแลระบบ
  const th = HtmlService.createTemplateFromFile("Home");
  th.email = email;
  th.bridgeUrl = BRIDGE_URL;
  th.canTeach = isTeacher_(email);
  th.canAdmin = isAdmin_(email);
  th.fullName = displayName_(email);

  return th
    .evaluate()
    .setTitle("ระบบเช็กชื่อ")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/* ===================== โทเค็นหมุนเวียน ===================== */

function slotNow_() {
  return Math.floor(Date.now() / (CONFIG.SLOT_SECONDS * 1000));
}

function sign_(sessionId, secret, slot) {
  const raw = Utilities.computeHmacSha256Signature(
    sessionId + ":" + slot,
    secret,
  );
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/, "").substring(0, 16);
}

function makeToken_(session, slot) {
  return slot.toString(36) + "." + sign_(session.id, session.secret, slot);
}

function verifyToken_(session, token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 2) return false;
  const slot = parseInt(parts[0], 36);
  if (isNaN(slot)) return false;
  const now = slotNow_();
  if (slot > now || now - slot > CONFIG.GRACE_SLOTS) return false;
  return sign_(session.id, session.secret, slot) === parts[1];
}

/* ===================== ตัวช่วยข้อมูล ===================== */

function db_() {
  const id = PropertiesService.getScriptProperties().getProperty("SS_ID");
  if (!id) throw new Error("ยังไม่ได้รัน setup()");
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  return db_().getSheetByName(name);
}

function rows_(name) {
  const sh = sheet_(name);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const head = HEADERS[name];
  return sh
    .getRange(2, 1, n - 1, head.length)
    .getValues()
    .map(function (r) {
      const o = {};
      head.forEach(function (h, i) {
        o[h] = r[i];
      });
      return o;
    });
}

function cacheGetJson_(key) {
  const v = CacheService.getScriptCache().get(key);
  return v ? JSON.parse(v) : null;
}

function cachePutJson_(key, obj, sec) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(obj), sec);
  } catch (err) {
    /* ข้อมูลใหญ่เกิน 100KB: ข้ามแคช */
  }
}

function students_() {
  let list = cacheGetJson_("students");
  if (!list) {
    list = rows_(SHEETS.STUDENTS)
      .filter(function (r) {
        return r.email;
      })
      .map(function (r) {
        return {
          email: String(r.email).trim().toLowerCase(),
          studentId: String(r.studentId),
          name: String(r.name),
          cls: String(r["class"]).trim(),
        };
      });
    cachePutJson_("students", list, CONFIG.CACHE_STUDENTS_SEC);
  }
  return list;
}

function isTeacher_(email) {
  let list = cacheGetJson_("teachers");
  if (!list) {
    list = rows_(SHEETS.TEACHERS).map(function (r) {
      return String(r.email).trim().toLowerCase();
    });
    cachePutJson_("teachers", list, 300);
  }
  return !!email && list.indexOf(email) >= 0;
}

function requireTeacher_() {
  const email = currentEmail_();
  if (!isTeacher_(email)) {
    throw new Error(
      "บัญชีนี้ไม่มีสิทธิ์อาจารย์ (" + (email || "อ่านอีเมลไม่ได้") + ")",
    );
  }
  return email;
}

/* ===================== Session ===================== */

function getSession_(id) {
  if (!id) return null;
  let s = cacheGetJson_("sess_" + id);
  if (s) return s;
  const sh = sheet_(SHEETS.SESSIONS);
  const cell = sh
    .getRange("A:A")
    .createTextFinder(id)
    .matchEntireCell(true)
    .findNext();
  if (!cell) return null;
  const r = sh.getRange(cell.getRow(), 1, 1, 7).getValues()[0];
  s = {
    id: String(r[0]),
    cls: String(r[1]),
    teacher: r[2],
    start: new Date(r[3]).getTime(),
    end: new Date(r[4]).getTime(),
    secret: String(r[5]),
    status: String(r[6]),
    row: cell.getRow(),
  };
  cachePutJson_("sess_" + id, s, CONFIG.CACHE_SESSION_SEC);
  return s;
}

function isOpen_(s) {
  return s.status === "open" && Date.now() < s.end;
}

/* ===================== ฝั่งอาจารย์ ===================== */

function getInit() {
  const email = requireTeacher_();
  const seen = {};
  students_().forEach(function (s) {
    seen[s.cls] = true;
  });
  return { email: email, classes: Object.keys(seen).sort() };
}

function startSession(cls, minutes) {
  const email = requireTeacher_();
  if (
    !students_().some(function (s) {
      return s.cls === cls;
    })
  )
    throw new Error("ไม่พบห้อง " + cls);
  const mins = Math.max(
    1,
    Math.min(CONFIG.MAX_MINUTES, Number(minutes) || CONFIG.DEFAULT_MINUTES),
  );
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const now = Date.now();
    const s = {
      id: Utilities.getUuid().replace(/-/g, "").substring(0, 10),
      cls: cls,
      teacher: email,
      start: now,
      end: now + mins * 60000,
      secret: Utilities.getUuid() + Utilities.getUuid(),
      status: "open",
      row: 0,
    };
    const sh = sheet_(SHEETS.SESSIONS);
    sh.appendRow([
      s.id,
      s.cls,
      s.teacher,
      new Date(s.start),
      new Date(s.end),
      s.secret,
      s.status,
    ]);
    s.row = sh.getLastRow();
    cachePutJson_("sess_" + s.id, s, CONFIG.CACHE_SESSION_SEC);
    cachePutJson_("att_" + s.id, [], CONFIG.CACHE_SESSION_SEC);
    return { id: s.id, cls: s.cls };
  } finally {
    lock.releaseLock();
  }
}

/** คืน URL ของ QR รอบปัจจุบัน + เวลาที่เหลือก่อนต้องเปลี่ยน */
function getToken(sessionId) {
  requireTeacher_();
  const s = getSession_(sessionId);
  if (!s || !isOpen_(s)) return { closed: true };
  const ms = CONFIG.SLOT_SECONDS * 1000;
  const slot = slotNow_();
  return {
    url:
      BRIDGE_URL +
      "?p=c&s=" +
      encodeURIComponent(s.id) +
      "&t=" +
      encodeURIComponent(makeToken_(s, slot)),
    msToNext: (slot + 1) * ms - Date.now(),
    slotMs: ms,
  };
}

function getStatus(sessionId) {
  requireTeacher_();
  const s = getSession_(sessionId);
  if (!s) throw new Error("ไม่พบรอบเช็กชื่อ");
  const roster = students_().filter(function (x) {
    return x.cls === s.cls;
  });
  const att = getAttendance_(s.id);
  const present = {};
  const devCount = {};
  att.forEach(function (a) {
    present[a.email] = true;
    if (a.deviceId) devCount[a.deviceId] = (devCount[a.deviceId] || 0) + 1;
  });
  return {
    open: isOpen_(s),
    cls: s.cls,
    total: roster.length,
    checked: att.length,
    remainingMs: Math.max(0, s.end - Date.now()),
    checkedList: att
      .slice()
      .sort(function (a, b) {
        return b.time - a.time;
      })
      .map(function (a) {
        return {
          email: a.email,
          name: a.name,
          studentId: a.studentId,
          time: a.time,
          method: a.method,
          shared: !!(a.deviceId && devCount[a.deviceId] > 1),
        };
      }),
    absentList: roster
      .filter(function (x) {
        return !present[x.email];
      })
      .map(function (x) {
        return { email: x.email, name: x.name, studentId: x.studentId };
      }),
  };
}

function stopSession(sessionId) {
  requireTeacher_();
  const s = getSession_(sessionId);
  if (!s) throw new Error("ไม่พบรอบเช็กชื่อ");
  s.status = "closed";
  s.end = Math.min(s.end, Date.now());
  const sh = sheet_(SHEETS.SESSIONS);
  sh.getRange(s.row, 5).setValue(new Date(s.end));
  sh.getRange(s.row, 7).setValue("closed");
  cachePutJson_("sess_" + s.id, s, CONFIG.CACHE_SESSION_SEC);
  return { ok: true };
}

/** อาจารย์ติ๊กให้เอง (นักเรียนไม่มีมือถือ / เน็ตมีปัญหา) */
function markManual(sessionId, email) {
  requireTeacher_();
  const s = getSession_(sessionId);
  if (!s) throw new Error("ไม่พบรอบเช็กชื่อ");
  const st = students_().filter(function (x) {
    return x.email === String(email).toLowerCase() && x.cls === s.cls;
  })[0];
  if (!st) throw new Error("ไม่พบนักเรียนในห้องนี้");
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const list = getAttendance_(s.id);
    if (
      !list.some(function (r) {
        return r.email === st.email;
      })
    )
      appendAttendance_(s, st, list, "", "manual", "");
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/** ยกเลิกการเช็กชื่อของนักเรียนคนหนึ่ง (กรณีตรวจพบการเช็กแทน) */
function removeAttendance(sessionId, email) {
  requireTeacher_();
  email = String(email).toLowerCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_(SHEETS.ATTENDANCE);
    const n = sh.getLastRow();
    if (n >= 2) {
      const vals = sh.getRange(2, 1, n - 1, 4).getValues(); // A..D = timestamp, sessionId, class, email
      for (let i = vals.length - 1; i >= 0; i--) {
        if (
          String(vals[i][1]) === sessionId &&
          String(vals[i][3]).toLowerCase() === email
        ) {
          sh.deleteRow(i + 2);
          break;
        }
      }
    }
    const list = getAttendance_(sessionId).filter(function (r) {
      return r.email !== email;
    });
    cachePutJson_("att_" + sessionId, list, CONFIG.CACHE_SESSION_SEC);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/* ===================== ฝั่งนักเรียน ===================== */

function checkin(sessionId, token, deviceId) {
  const email = currentEmail_();
  if (!email) {
    return fail_("NO_LOGIN", "อ่านอีเมลไม่ได้ กรุณาล็อกอินด้วยบัญชี Google");
  }
  const s = getSession_(sessionId);
  if (!s) return fail_("NO_SESSION", "ไม่พบรอบเช็กชื่อนี้");
  if (!isOpen_(s)) return fail_("CLOSED", "ปิดรับเช็กชื่อแล้ว");
  if (!verifyToken_(s, token))
    return fail_("EXPIRED", "QR หมดอายุ กรุณาสแกนใหม่จากหน้าจอหน้าห้อง");
  const st = students_().filter(function (x) {
    return x.email === email;
  })[0];
  if (!st)
    return fail_(
      "NOT_STUDENT",
      "ไม่พบอีเมลนี้ในรายชื่อนักเรียน (" + email + ")",
    );
  if (st.cls !== s.cls)
    return fail_("WRONG_CLASS", "บัญชีนี้ไม่ได้อยู่ในห้อง " + s.cls);

  deviceId = String(deviceId || "").substring(0, 64);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const list = getAttendance_(s.id);
    const done = list.filter(function (r) {
      return r.email === email;
    })[0];
    if (done) return okResult_(st, done.time, true);
    const shared =
      deviceId &&
      list.some(function (r) {
        return r.deviceId === deviceId && r.email !== email;
      });
    const rec = appendAttendance_(
      s,
      st,
      list,
      deviceId,
      "qr",
      shared ? "DEVICE_SHARED" : "",
    );
    return okResult_(st, rec.time, false);
  } finally {
    lock.releaseLock();
  }
}

/* ===================== บันทึกการเช็กชื่อ ===================== */

function getAttendance_(sessionId) {
  let list = cacheGetJson_("att_" + sessionId);
  if (list) return list;
  list = rows_(SHEETS.ATTENDANCE)
    .filter(function (r) {
      return String(r.sessionId) === String(sessionId);
    })
    .map(function (r) {
      return {
        email: String(r.email).toLowerCase(),
        studentId: String(r.studentId),
        name: String(r.name),
        time: new Date(r.timestamp).getTime(),
        deviceId: String(r.deviceId || ""),
        method: String(r.method),
        flag: String(r.flag || ""),
      };
    });
  cachePutJson_("att_" + sessionId, list, CONFIG.CACHE_SESSION_SEC);
  return list;
}

function appendAttendance_(s, st, list, deviceId, method, flag) {
  const rec = {
    email: st.email,
    studentId: st.studentId,
    name: st.name,
    time: Date.now(),
    deviceId: deviceId,
    method: method,
    flag: flag,
  };
  sheet_(SHEETS.ATTENDANCE).appendRow([
    new Date(rec.time),
    s.id,
    s.cls,
    st.email,
    st.studentId,
    st.name,
    deviceId,
    method,
    flag,
  ]);
  list.push(rec);
  cachePutJson_("att_" + s.id, list, CONFIG.CACHE_SESSION_SEC);
  return rec;
}

/* ===================== ผู้ดูแลระบบ: รายงานและปรับสถานะ ===================== */

function isAdmin_(email) {
  let list = cacheGetJson_("admins");
  if (!list) {
    list = rows_(SHEETS.ADMINS).map(function (r) {
      return String(r.email).trim().toLowerCase();
    });
    cachePutJson_("admins", list, 300);
  }
  return !!email && list.indexOf(email) >= 0;
}

function requireAdmin_() {
  const email = currentEmail_();
  if (!isAdmin_(email)) throw new Error("เฉพาะผู้ดูแลระบบเท่านั้น");
  return email;
}

function getAdminInit() {
  const email = requireAdmin_();
  const seen = {};
  students_().forEach(function (s) {
    seen[s.cls] = true;
  });
  return { email: email, classes: Object.keys(seen).sort() };
}

/**
 * รายงานของห้องหนึ่ง ในช่วงวันที่ from..to (รูปแบบ yyyy-MM-dd เว้นว่างได้)
 * คืนนักเรียนทั้งห้อง + รายการรอบที่เสร็จแล้ว แต่ละรอบมีสตริง st ยาวเท่าจำนวนนักเรียน
 * (ตัวอักษรต่อคน: p มา, l สาย, v ลา, a ขาด) และ ad (1 = ถูกแอดมินปรับ) กับ notes
 * ลำดับสถานะ: ที่แอดมินปรับ > มีบันทึกเช็กชื่อ = มา > ไม่มี = ขาด
 * รอบที่ยังเปิดอยู่และรอบที่ถูกยกเลิก (void) จะไม่ถูกนับ
 */
function getReport(cls, from, to) {
  requireAdmin_();
  const roster = students_()
    .filter(function (x) {
      return x.cls === cls;
    })
    .sort(function (a, b) {
      return cmpId_(a.studentId, b.studentId);
    });
  const now = Date.now();
  const fmt = function (ms, f) {
    return Utilities.formatDate(new Date(ms), CONFIG.TZ, f);
  };

  const sess = rows_(SHEETS.SESSIONS)
    .filter(function (r) {
      if (String(r["class"]) !== cls) return false;
      const status = String(r.status);
      if (status === "void") return false;
      if (status === "open" && now < new Date(r.end).getTime()) return false;
      const d = fmt(new Date(r.start).getTime(), "yyyy-MM-dd");
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    })
    .sort(function (a, b) {
      return new Date(a.start).getTime() - new Date(b.start).getTime();
    });

  const ids = {};
  sess.forEach(function (r) {
    ids[String(r.sessionId)] = true;
  });

  const att = {};
  const sh = sheet_(SHEETS.ATTENDANCE);
  const n = sh.getLastRow();
  if (n >= 2 && sess.length) {
    sh.getRange(2, 1, n - 1, 8)
      .getValues()
      .forEach(function (r) {
        // A = timestamp, B..D = sessionId, class, email, H = method
        const sid = String(r[1]);
        if (!ids[sid]) return;
        (att[sid] = att[sid] || {})[String(r[3]).toLowerCase()] = {
          t: new Date(r[0]).getTime(),
          m: String(r[7] || ""),
        };
      });
  }

  const adj = {};
  rows_(SHEETS.ADJUSTMENTS).forEach(function (r) {
    adj[String(r.sessionId) + "|" + String(r.email).toLowerCase()] = {
      code: STATUS_CODE[String(r.status)] || "a",
      note: String(r.note || ""),
    };
  });

  // ตารางสอน: ใช้หาชื่อวิชา/คาบที่ ของแต่ละรอบ (จับคู่ด้วย อาจารย์ + ห้อง + วัน + เวลาเริ่ม)
  const sched = rows_(SHEETS.SCHEDULE)
    .filter(function (r) {
      return String(r["class"]) === cls;
    })
    .map(function (r) {
      return {
        teacher: String(r.teacher).trim().toLowerCase(),
        day: Number(r.day),
        period: Number(r.period),
        subject: String(r.subject || ""),
        room: String(r.room || ""),
      };
    });

  return {
    cls: cls,
    students: roster.map(function (x) {
      return { email: x.email, studentId: x.studentId, name: x.name };
    }),
    sessions: sess.map(function (r) {
      const sid = String(r.sessionId);
      const startMs = new Date(r.start).getTime();
      let st = "",
        ad = "";
      const notes = {};
      const ts = {}; // เวลาที่นักเรียนเช็กชื่อจริง (HH:mm:ss) จากชีต Attendance
      roster.forEach(function (x) {
        const a = adj[sid + "|" + x.email];
        const rec = att[sid] && att[sid][x.email];
        if (rec) ts[x.email] = { t: fmt(rec.t, "HH:mm:ss"), m: rec.m };
        if (a) {
          st += a.code;
          ad += "1";
          if (a.note) notes[x.email] = a.note;
        } else {
          st += rec ? "p" : "a";
          ad += "0";
        }
      });
      const slot = matchSlot_(
        sched,
        String(r.teacher).trim().toLowerCase(),
        startMs,
      );
      return {
        id: sid,
        date: fmt(startMs, "yyyy-MM-dd"),
        time: fmt(startMs, "HH:mm"),
        st: st,
        ad: ad,
        notes: notes,
        ts: ts,
        subject: slot ? slot.subject : "",
        period: slot ? slot.period : 0,
        room: slot ? slot.room : "",
      };
    }),
  };
}

// เวลาคาบเรียน (ต้องตรงกับตาราง PERIODS ใน Admin.html)
const PERIOD_TIMES = [
  [1, "08:30", "09:20"],
  [2, "09:20", "10:10"],
  [3, "10:20", "11:10"],
  [4, "11:10", "12:00"],
  [5, "13:00", "13:50"],
  [6, "13:50", "14:40"],
  [7, "14:50", "15:40"],
  [8, "15:40", "16:30"],
];

/** หาช่องตารางสอนที่ตรงกับรอบเช็กชื่อ (อาจารย์เดียวกัน วันเดียวกัน คาบที่เวลาเริ่มอยู่ในช่วง หรือเปิดก่อนคาบไม่เกิน 15 นาที) */
function matchSlot_(sched, teacher, startMs) {
  const dow = Number(Utilities.formatDate(new Date(startMs), CONFIG.TZ, "u")); // 1=จันทร์..7=อาทิตย์
  const hm = Utilities.formatDate(new Date(startMs), CONFIG.TZ, "HH:mm").split(
    ":",
  );
  const m = Number(hm[0]) * 60 + Number(hm[1]);
  const toMin = function (t) {
    const a = t.split(":");
    return Number(a[0]) * 60 + Number(a[1]);
  };
  let period = 0;
  for (let i = 0; i < PERIOD_TIMES.length && !period; i++) {
    if (m >= toMin(PERIOD_TIMES[i][1]) && m < toMin(PERIOD_TIMES[i][2]))
      period = PERIOD_TIMES[i][0];
  }
  for (let i = 0; i < PERIOD_TIMES.length && !period; i++) {
    const d = toMin(PERIOD_TIMES[i][1]) - m;
    if (d > 0 && d <= 15) period = PERIOD_TIMES[i][0];
  }
  if (!period) return null;
  return (
    sched.filter(function (x) {
      return x.teacher === teacher && x.day === dow && x.period === period;
    })[0] || { subject: "", period: period, room: "" }
  );
}

/**
 * ตั้งสถานะของนักเรียนในรอบหนึ่ง: present | late | leave | absent | auto (ล้างที่ปรับ กลับไปตามระบบ)
 * คืนรหัสสถานะที่มีผลตอนนี้
 */
function setStatus(sessionId, email, status, note) {
  const admin = requireAdmin_();
  email = String(email).toLowerCase();
  const s = getSession_(sessionId);
  if (!s) throw new Error("ไม่พบรอบเช็กชื่อ");
  if (
    !students_().some(function (x) {
      return x.email === email && x.cls === s.cls;
    })
  ) {
    throw new Error("ไม่พบนักเรียนในห้องนี้");
  }
  if (status !== "auto" && !STATUS_CODE[status])
    throw new Error("สถานะไม่ถูกต้อง");
  note = String(note || "").substring(0, 200);
  if (/^[=+\-@]/.test(note)) note = "'" + note; // กันสูตรในชีต

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_(SHEETS.ADJUSTMENTS);
    const n = sh.getLastRow();
    let row = 0;
    if (n >= 2) {
      const keys = sh.getRange(2, 1, n - 1, 2).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (
          String(keys[i][0]) === sessionId &&
          String(keys[i][1]).toLowerCase() === email
        ) {
          row = i + 2;
          break;
        }
      }
    }
    if (status === "auto") {
      if (row) sh.deleteRow(row);
      const had = getAttendance_(sessionId).some(function (r) {
        return r.email === email;
      });
      return { code: had ? "p" : "a", adjusted: false };
    }
    const rec = [sessionId, email, status, note, admin, new Date()];
    if (row) sh.getRange(row, 1, 1, 6).setValues([rec]);
    else sh.appendRow(rec);
    return { code: STATUS_CODE[status], adjusted: true };
  } finally {
    lock.releaseLock();
  }
}

/** ยกเลิกรอบเช็กชื่อ (เช่น อาจารย์กดเริ่มผิด) รอบนี้จะไม่ถูกนับในรายงาน ข้อมูลยังอยู่ในชีต */
function voidSession(sessionId) {
  requireAdmin_();
  const s = getSession_(sessionId);
  if (!s) throw new Error("ไม่พบรอบเช็กชื่อ");
  s.status = "void";
  sheet_(SHEETS.SESSIONS).getRange(s.row, 7).setValue("void");
  cachePutJson_("sess_" + s.id, s, CONFIG.CACHE_SESSION_SEC);
  return { ok: true };
}

function fail_(code, msg) {
  return { ok: false, code: code, msg: msg };
}

function okResult_(st, timeMs, already) {
  return {
    ok: true,
    already: already,
    name: st.name,
    studentId: st.studentId,
    cls: st.cls,
    timeText: Utilities.formatDate(new Date(timeMs), CONFIG.TZ, "HH:mm:ss"),
  };
}

/* ===================== ตารางสอน & แดชบอร์ด ===================== */
const SCHEDULE_CFG = { DAYS: 5, PERIODS: 8, LOW_RATE: 80 }; // วันที่ 1-5 = จันทร์-ศุกร์

function displayName_(email) {
  const key = String(email || "")
    .trim()
    .toLowerCase();
  const pick = function (sheetName) {
    const hit = rows_(sheetName).filter(function (r) {
      return String(r.email).trim().toLowerCase() === key;
    })[0];
    return hit ? String(hit.fullName || "").trim() : "";
  };
  return pick(SHEETS.TEACHERS) || pick(SHEETS.ADMINS) || "";
}

function mySchedule_(email) {
  return rows_(SHEETS.SCHEDULE)
    .filter(function (r) {
      return String(r.teacher).trim().toLowerCase() === email;
    })
    .map(function (r) {
      return {
        day: Number(r.day),
        period: Number(r.period),
        cls: String(r["class"] || ""),
        subject: String(r.subject || ""),
        room: String(r.room || ""),
      };
    });
}

// ลบช่องตารางสอนของอาจารย์ในวัน/ช่วงคาบที่ระบุ (เรียกภายใน lock เท่านั้น)
function removeSlots_(email, day, fromP, toP) {
  const sh = sheet_(SHEETS.SCHEDULE);
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const vals = sh.getRange(2, 2, n - 1, 3).getValues(); // B..D = teacher, day, period
  const hit = [];
  vals.forEach(function (r, i) {
    const p = Number(r[2]);
    if (
      String(r[0]).trim().toLowerCase() === email &&
      Number(r[1]) === day &&
      p >= fromP &&
      p <= toP
    )
      hit.push(i + 2);
  });
  for (let k = hit.length - 1; k >= 0; k--) sh.deleteRow(hit[k]);
  return hit.length;
}

function saveSchedule(day, fromP, toP, cls, subject, room) {
  const email = String(requireAdmin_()).toLowerCase();
  day = Number(day);
  fromP = Number(fromP);
  toP = Number(toP) || fromP;
  if (!(day >= 1 && day <= SCHEDULE_CFG.DAYS)) throw new Error("วันไม่ถูกต้อง");
  if (!(fromP >= 1 && toP >= fromP && toP <= SCHEDULE_CFG.PERIODS))
    throw new Error("คาบไม่ถูกต้อง");
  subject = String(subject || "")
    .trim()
    .substring(0, 80);
  cls = String(cls || "")
    .trim()
    .substring(0, 40);
  room = String(room || "")
    .trim()
    .substring(0, 40);
  if (!subject) throw new Error("กรุณาใส่ชื่อวิชา");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    removeSlots_(email, day, fromP, toP); // แทนที่ช่องเดิมในช่วงที่เลือก
    const sh = sheet_(SHEETS.SCHEDULE);
    const now = new Date();
    const out = [];
    for (let p = fromP; p <= toP; p++) {
      out.push([
        Utilities.getUuid().replace(/-/g, "").substring(0, 10),
        email,
        day,
        p,
        cls,
        subject,
        room,
        now,
      ]);
    }
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 5, out.length, 3).setNumberFormat("@");
    sh.getRange(start, 1, out.length, out[0].length).setValues(out);
  } finally {
    lock.releaseLock();
  }
  return { schedule: mySchedule_(email) };
}

function deleteSchedule(day, period) {
  const email = String(requireAdmin_()).toLowerCase();
  day = Number(day);
  period = Number(period);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    removeSlots_(email, day, period, period);
  } finally {
    lock.releaseLock();
  }
  return { schedule: mySchedule_(email) };
}

// สรุปการเข้าเรียนทุกห้องในช่วงวันที่ (yyyy-MM-dd) ใช้ตรรกะเดียวกับ getReport
function summary_(from, to) {
  const re = /^\d{4}-\d{2}-\d{2}$/;
  from = re.test(from || "") ? from : "";
  to = re.test(to || "") ? to : "";
  const now = Date.now();
  const ymd = function (v) {
    return Utilities.formatDate(new Date(v), CONFIG.TZ, "yyyy-MM-dd");
  };
  const zero = function () {
    return { p: 0, l: 0, v: 0, a: 0 };
  };

  const sess = rows_(SHEETS.SESSIONS).filter(function (r) {
    const status = String(r.status);
    if (status === "void") return false;
    if (status === "open" && now < new Date(r.end).getTime()) return false;
    const d = ymd(r.start);
    if (from && d < from) return false;
    if (to && d > to) return false;
    return true;
  });

  const ids = {};
  sess.forEach(function (r) {
    ids[String(r.sessionId)] = true;
  });

  const att = {};
  const sh = sheet_(SHEETS.ATTENDANCE);
  const n = sh.getLastRow();
  if (n >= 2 && sess.length) {
    sh.getRange(2, 2, n - 1, 3)
      .getValues()
      .forEach(function (r) {
        const sid = String(r[0]);
        if (!ids[sid]) return;
        (att[sid] = att[sid] || {})[String(r[2]).toLowerCase()] = true;
      });
  }

  const adj = {};
  rows_(SHEETS.ADJUSTMENTS).forEach(function (r) {
    adj[String(r.sessionId) + "|" + String(r.email).toLowerCase()] =
      STATUS_CODE[String(r.status)] || "a";
  });

  const roster = {};
  students_().forEach(function (s) {
    (roster[s.cls] = roster[s.cls] || []).push(s);
  });

  const all = zero();
  const byClass = {};
  const byStu = {};
  sess.forEach(function (r) {
    const sid = String(r.sessionId);
    const cls = String(r["class"]);
    const bc = (byClass[cls] = byClass[cls] || {
      cls: cls,
      sessions: 0,
      c: zero(),
    });
    bc.sessions++;
    (roster[cls] || []).forEach(function (x) {
      const em = String(x.email).toLowerCase();
      const code =
        adj[sid + "|" + em] || (att[sid] && att[sid][em] ? "p" : "a");
      all[code]++;
      bc.c[code]++;
      const key = cls + "|" + em;
      const st = (byStu[key] = byStu[key] || {
        cls: cls,
        studentId: x.studentId,
        name: x.name,
        c: zero(),
      });
      st.c[code]++;
    });
  });

  const rate = function (c) {
    const t = c.p + c.l + c.v + c.a;
    return t ? ((c.p + c.l) / t) * 100 : 0;
  };
  const watch = Object.keys(byStu)
    .map(function (k) {
      return byStu[k];
    })
    .filter(function (s) {
      return (
        s.c.p + s.c.l + s.c.v + s.c.a >= 3 && rate(s.c) < SCHEDULE_CFG.LOW_RATE
      );
    })
    .sort(function (a, b) {
      return rate(a.c) - rate(b.c);
    })
    .slice(0, 8);

  return {
    sessions: sess.length,
    counts: all,
    lowRate: SCHEDULE_CFG.LOW_RATE,
    byClass: Object.keys(byClass)
      .sort()
      .map(function (k) {
        return byClass[k];
      }),
    watch: watch,
  };
}

function getDashboard(from, to) {
  const email = String(requireAdmin_()).toLowerCase();
  return {
    email: email,
    name: displayName_(email),
    schedule: mySchedule_(email),
    summary: summary_(from, to),
    bridgeUrl: BRIDGE_URL, // ใหม่
  };
}

/* ===================== จัดการรายชื่อนักเรียน ===================== */

// เทียบรหัส/เลขที่แบบ "ตัวเลขมาก่อน": 2 มาก่อน 10 (ถ้าไม่ใช่ตัวเลขล้วนใช้เทียบตัวอักษร)
function cmpId_(a, b) {
  a = String(a);
  b = String(b);
  const na = Number(a),
    nb = Number(b);
  if (a !== "" && b !== "" && !isNaN(na) && !isNaN(nb) && na !== nb)
    return na - nb;
  return a < b ? -1 : a > b ? 1 : 0;
}

// เขียนแถวนักเรียนลงชีตแบบข้อความล้วน (รักษาเลขศูนย์นำหน้า และกันข้อความที่ขึ้นต้นด้วย = ถูกตีเป็นสูตร)
function writeStudentRow_(sh, row, email, studentId, name, cls) {
  const rng = sh.getRange(row, 1, 1, 4);
  rng.setNumberFormat("@");
  rng.setValues([[email, studentId, name, cls]]);
}

// เปลี่ยนอีเมลเดิมเป็นอีเมลใหม่ในคอลัมน์ที่ระบุ เพื่อให้ประวัติเช็กชื่อ/การปรับสถานะยังติดตัวนักเรียน
function renameEmail_(sheetName, col, oldEmail, newEmail) {
  const sh = sheet_(sheetName);
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const rng = sh.getRange(2, col, n - 1, 1);
  const vals = rng.getValues();
  let changed = 0;
  vals.forEach(function (r) {
    if (String(r[0]).trim().toLowerCase() === oldEmail) {
      r[0] = newEmail;
      changed++;
    }
  });
  if (changed) rng.setValues(vals);
  return changed;
}

function getStudentsPage(cls) {
  requireAdmin_();
  cls = String(cls || "").trim();
  const all = students_(); // ใช้ฟังก์ชันแคชเดิมของระบบ
  const seenCls = {};
  all.forEach(function (s) {
    if (s.cls) seenCls[s.cls] = true;
  });
  const classes = Object.keys(seenCls).sort();

  // ถ้าห้องที่ขอไม่มีแล้ว (เช่น ลบคนสุดท้ายของห้อง) ให้ไปห้องแรกแทน
  const targetCls = seenCls[cls] ? cls : classes[0] || "";
  const list = all
    .filter(function (s) {
      return !targetCls || s.cls === targetCls;
    })
    .sort(function (a, b) {
      return cmpId_(a.studentId, b.studentId);
    });

  return {
    classes: classes,
    currentClass: targetCls,
    students: list,
  };
}

//เพิ่มรายชื่อนักเรียน
function addStudent(studentId, name, cls, email) {
  requireAdmin_();
  studentId = String(studentId || "").trim();
  name = String(name || "").trim();
  cls = String(cls || "").trim();
  email = String(email || "")
    .trim()
    .toLowerCase();

  if (!studentId) throw new Error("กรุณากรอกเลขที่");
  if (!name) throw new Error("กรุณากรอกชื่อ-นามสกุล");
  if (!cls) throw new Error("กรุณาระบุห้องเรียน");
  if (!email || email.indexOf("@") === -1) throw new Error("อีเมลไม่ถูกต้อง");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_(SHEETS.STUDENTS);
    const rows = rows_(SHEETS.STUDENTS);

    // อีเมลต้องไม่ซ้ำทั้งระบบ / รหัส-เลขที่ต้องไม่ซ้ำภายในห้องเดียวกัน
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i].email).trim().toLowerCase() === email) {
        throw new Error("อีเมล '" + email + "' มีอยู่ในระบบแล้ว");
      }
      if (
        String(rows[i].studentId).trim() === studentId &&
        String(rows[i]["class"]).trim() === cls
      ) {
        throw new Error(
          "เลขที่ '" + studentId + "' มีอยู่ในห้อง " + cls + " แล้ว",
        );
      }
    }

    // Headers ของ Students: ["email", "studentId", "name", "class"]
    writeStudentRow_(sh, sh.getLastRow() + 1, email, studentId, name, cls);
    clearCache(); // ล้างแคชเพื่อให้ระบบโหลดข้อมูลนักเรียนใหม่ทันที
  } finally {
    lock.releaseLock();
  }
  return getStudentsPage(cls);
}

//แก้ไขรายชื่อนักเรียน
function updateStudent(oldEmail, newStudentId, newName, newClass, newEmail) {
  requireAdmin_();
  oldEmail = String(oldEmail || "")
    .trim()
    .toLowerCase();
  newStudentId = String(newStudentId || "").trim();
  newName = String(newName || "").trim();
  newClass = String(newClass || "").trim();
  newEmail = String(newEmail || "")
    .trim()
    .toLowerCase();

  if (!oldEmail) throw new Error("ไม่พบอีเมลเดิมของนักเรียน");
  if (!newStudentId) throw new Error("กรุณากรอกเลขที่");
  if (!newName) throw new Error("กรุณากรอกชื่อ-นามสกุล");
  if (!newClass) throw new Error("กรุณาระบุห้องเรียน");
  if (!newEmail || newEmail.indexOf("@") === -1)
    throw new Error("อีเมลไม่ถูกต้อง");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_(SHEETS.STUDENTS);
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const emailIdx = headers.indexOf("email");
    const idIdx = headers.indexOf("studentId");
    const classIdx = headers.indexOf("class");

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][emailIdx]).trim().toLowerCase();
      const rowId = String(data[i][idIdx]).trim();
      const rowCls = String(data[i][classIdx]).trim();

      if (rowEmail === oldEmail) {
        rowIndex = i + 1; // แถวใน Google Sheet เริ่มที่ 1
        continue;
      }
      // ซ้ำกับคนอื่น: อีเมลซ้ำทั้งระบบ / รหัส-เลขที่ซ้ำภายในห้องเดียวกัน
      if (rowEmail === newEmail)
        throw new Error("อีเมล '" + newEmail + "' มีผู้ใช้งานแล้ว");
      if (rowId === newStudentId && rowCls === newClass)
        throw new Error(
          "เลขที่ '" +
            newStudentId +
            "' มีผู้ใช้งานในห้อง " +
            newClass +
            " แล้ว",
        );
    }

    if (rowIndex === -1) throw new Error("ไม่พบข้อมูลนักเรียนในระบบ");

    // อัปเดตข้อมูลลงชีต (Headers: email, studentId, name, class)
    writeStudentRow_(sh, rowIndex, newEmail, newStudentId, newName, newClass);

    // เปลี่ยนอีเมล: ย้ายประวัติเช็กชื่อและการปรับสถานะตามไปด้วย ไม่เช่นนั้นจะนับเป็น "ขาด" ย้อนหลัง
    if (newEmail !== oldEmail) {
      renameEmail_(SHEETS.ATTENDANCE, 4, oldEmail, newEmail); // D = email
      renameEmail_(SHEETS.ADJUSTMENTS, 2, oldEmail, newEmail); // B = email
    }

    clearCache();
  } finally {
    lock.releaseLock();
  }
  return getStudentsPage(newClass);
}

//ลบรายชื่อนักเรียน
function deleteStudent(email, cls) {
  requireAdmin_();
  email = String(email || "")
    .trim()
    .toLowerCase();
  cls = String(cls || "").trim();
  if (!email) throw new Error("ไม่พบอีเมลนักเรียน");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_(SHEETS.STUDENTS);
    const data = sh.getDataRange().getValues();
    const emailIdx = data[0].indexOf("email");

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailIdx]).trim().toLowerCase() === email) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw new Error("ไม่พบข้อมูลนักเรียนในระบบ");
    sh.deleteRow(rowIndex);
    clearCache();
  } finally {
    lock.releaseLock();
  }
  return getStudentsPage(cls);
}

function downloadExcelReport(cls, from, to) {
  requireAdmin_();
  const report = getReport(cls, from, to);
  const sessions = report.sessions;
  const students = report.students;

  const THAI_MONTHS = [
    "มกราคม",
    "กุมภาพันธ์",
    "มีนาคม",
    "เมษายน",
    "พฤษภาคม",
    "มิถุนายน",
    "กรกฎาคม",
    "สิงหาคม",
    "กันยายน",
    "ตุลาคม",
    "พฤศจิกายน",
    "ธันวาคม",
  ];
  const DOW = ["อา.", "จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส."];

  // แยกปี/เดือน/วัน จากข้อความ yyyy-MM-dd โดยตรง (ไม่พึ่ง locale ของเซิร์ฟเวอร์)
  function parseDate_(str) {
    const p = String(str).split("-");
    const y = +p[0],
      m = +p[1],
      d = +p[2];
    return {
      y: y,
      m: m,
      d: d,
      dow: new Date(Date.UTC(y, m - 1, d)).getUTCDay(),
    };
  }

  // ---------- จัดกลุ่ม: เดือน -> วัน -> คาบ (เหมือนหน้าเว็บ) ----------
  const monthMap = {};
  const monthOrder = [];
  sessions.forEach(function (se) {
    const dt = parseDate_(se.date);
    const mKey = se.date.substring(0, 7);
    if (!monthMap[mKey]) {
      monthMap[mKey] = {
        name: THAI_MONTHS[dt.m - 1] + " " + (dt.y + 543),
        days: {},
        dayOrder: [],
        count: 0,
      };
      monthOrder.push(mKey);
    }
    const m = monthMap[mKey];
    if (!m.days[se.date]) {
      m.days[se.date] = { dt: dt, sessions: [] };
      m.dayOrder.push(se.date);
    }
    m.days[se.date].sessions.push(se);
    m.count++;
  });

  // ---------- ลิสต์คอลัมน์แบบแบน (ใช้ร่วมกันทั้ง header และข้อมูล) ----------
  // ms/me = ขอบซ้าย/ขวาของเดือน (เข้ม), ds/de = ขอบซ้าย/ขวาของวัน (จางกว่า)
  const cols = [];
  monthOrder.forEach(function (mKey) {
    const m = monthMap[mKey];
    m.dayOrder.forEach(function (dateKey, dIdx) {
      const items = m.days[dateKey].sessions;
      items.forEach(function (se, sIdx) {
        const b = [];
        if (sIdx === 0) b.push(dIdx === 0 ? "ms" : "ds");
        if (sIdx === items.length - 1)
          b.push(dIdx === m.dayOrder.length - 1 ? "me" : "de");
        cols.push({ se: se, border: b.join(" ") });
      });
    });
  });

  // ---------- ตำแหน่งคอลัมน์ / แถว ----------
  function colLetter(colIdx) {
    let temp,
      letter = "";
    while (colIdx >= 0) {
      temp = colIdx % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      colIdx = Math.floor(colIdx / 26) - 1;
    }
    return letter;
  }

  const startRow = 4; // แถว 1-3 = หัวตาราง, นักเรียนเริ่มแถว 4
  const startColIdx = 3; // คอลัมน์ D
  const endColIdx = startColIdx + cols.length - 1;
  const startColLet = colLetter(startColIdx);
  const endColLet = colLetter(endColIdx);

  // ลำดับคอลัมน์สรุป ตรงกับหน้าเว็บ: ทั้งหมด, มา, สาย, ขาด, ลา, รวมไม่มา, ร้อยละ
  const totalCol = colLetter(endColIdx + 1);
  const presentCol = colLetter(endColIdx + 2);
  const lateCol = colLetter(endColIdx + 3);
  const absentCol = colLetter(endColIdx + 4);
  const leaveCol = colLetter(endColIdx + 5);
  const notComeCol = colLetter(endColIdx + 6);

  // ---------- CSS ----------
  let html =
    '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">\n<head>\n<meta http-equiv="content-type" content="text/html; charset=UTF-8">\n<style>\n' +
    "table { border-collapse: collapse; font-family: sans-serif; font-size: 11pt; }\n" +
    "th, td { border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; vertical-align: middle; }\n" +
    "th { background-color: #2563eb; color: #ffffff; font-weight: bold; }\n" +
    "th.h1 { background-color: #1e3a8a; }\n" +
    "th.h3 { background-color: #dbeafe; color: #1e3a8a; font-size: 10pt; }\n" +
    ".name { text-align: left; }\n" +
    "td.s { border: 1px solid #eceff7; }\n" +
    "td.p { color: #16a34a; }\n" +
    "td.l { color: #d97706; font-weight: bold; }\n" +
    "td.v { color: #2563eb; font-weight: bold; }\n" +
    "td.a { color: #dc2626; font-weight: bold; }\n" +
    "td.ds { border-left: 1px solid #94a3b8; }\n" +
    "td.de { border-right: 1px solid #94a3b8; }\n" +
    "td.ms { border-left: 2px solid #475569; }\n" +
    "td.me { border-right: 2px solid #475569; }\n" +
    "</style>\n</head>\n<body>\n<table>\n";

  // ความกว้างคอลัมน์
  html += '<colgroup><col width="40"><col width="60"><col width="200">';
  cols.forEach(function () {
    html += '<col width="38">';
  });
  html +=
    '<col width="55"><col width="45"><col width="45"><col width="45"><col width="45"><col width="65"><col width="60"></colgroup>\n';

  // ---------- Header แถว 1: เดือน ----------
  html +=
    '<tr>\n<th rowspan="3" class="h1">ลำดับ</th>\n<th rowspan="3" class="h1">เลขที่</th>\n<th rowspan="3" class="h1">ชื่อ-สกุล</th>\n';
  monthOrder.forEach(function (mKey) {
    const m = monthMap[mKey];
    html +=
      '<th colspan="' + m.count + '" class="h1">' + escXml_(m.name) + "</th>\n";
  });
  ["ทั้งหมด", "มา", "สาย", "ขาด", "ลา", "รวมไม่มา", "ร้อยละ"].forEach(
    function (t) {
      html += '<th rowspan="3" class="h1">' + t + "</th>\n";
    },
  );
  html += "</tr>\n";

  // ---------- Header แถว 2: วัน (คลุมทุกคาบของวันนั้น) ----------
  html += "<tr>\n";
  monthOrder.forEach(function (mKey) {
    const m = monthMap[mKey];
    m.dayOrder.forEach(function (dateKey) {
      const day = m.days[dateKey];
      html +=
        '<th colspan="' +
        day.sessions.length +
        '">' +
        DOW[day.dt.dow] +
        " " +
        day.dt.d +
        "</th>\n";
    });
  });
  html += "</tr>\n";

  // ---------- Header แถว 3: คาบ ----------
  html += "<tr>\n";
  cols.forEach(function (col) {
    const se = col.se;
    const label = se.period ? "คาบ " + se.period : se.time || "";
    html += '<th class="h3">' + escXml_(String(label)) + "</th>\n";
  });
  html += "</tr>\n";

  // ---------- แถวนักเรียน ----------
  students.forEach(function (s, idx) {
    const r = startRow + idx;
    const range = startColLet + r + ":" + endColLet + r;

    html += "<tr>\n";
    html += "<td>" + (idx + 1) + "</td>\n";
    html += "<td>" + escXml_(s.studentId) + "</td>\n";
    html += '<td class="name">' + escXml_(s.name) + "</td>\n";

    cols.forEach(function (col) {
      const code = col.se.st[idx];
      let sym = "";
      if (code === "p") sym = "/";
      else if (code === "l") sym = "ส";
      else if (code === "v") sym = "ล";
      else if (code === "a") sym = "ข";
      const cssCls =
        "s" + (sym ? " " + code : "") + (col.border ? " " + col.border : "");
      html += '<td class="' + cssCls + '">' + sym + "</td>\n";
    });

    // สูตรสรุปผล (ลำดับตรงกับหน้าเว็บ)
    html += "<td>=COUNTA(" + range + ")</td>\n"; // ทั้งหมด
    html += "<td>=COUNTIF(" + range + ',"/")</td>\n'; // มา
    html += "<td>=COUNTIF(" + range + ',"ส")</td>\n'; // สาย
    html += "<td>=COUNTIF(" + range + ',"ข")</td>\n'; // ขาด
    html += "<td>=COUNTIF(" + range + ',"ล")</td>\n'; // ลา
    html += "<td>=" + leaveCol + r + "+" + absentCol + r + "</td>\n"; // รวมไม่มา = ลา + ขาด
    html +=
      "<td>=IF(" +
      totalCol +
      r +
      "&gt;0,ROUND((" +
      presentCol +
      r +
      "+" +
      lateCol +
      r +
      ")/" +
      totalCol +
      r +
      "*100,1),0)</td>\n"; // ร้อยละ
    html += "</tr>\n";
  });

  html += "</table>\n</body>\n</html>";

  return {
    filename: "attendance_" + cls.replace(/[^\wก-๙.-]+/g, "_") + ".xls",
    content: Utilities.base64Encode(html, Utilities.Charset.UTF_8),
  };
}

function escXml_(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
