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
    sh.getRange(2, 2, n - 1, 3)
      .getValues()
      .forEach(function (r) {
        // B..D = sessionId, class, email
        const sid = String(r[0]);
        if (!ids[sid]) return;
        (att[sid] = att[sid] || {})[String(r[2]).toLowerCase()] = true;
      });
  }

  const adj = {};
  rows_(SHEETS.ADJUSTMENTS).forEach(function (r) {
    adj[String(r.sessionId) + "|" + String(r.email).toLowerCase()] = {
      code: STATUS_CODE[String(r.status)] || "a",
      note: String(r.note || ""),
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
      roster.forEach(function (x) {
        const a = adj[sid + "|" + x.email];
        if (a) {
          st += a.code;
          ad += "1";
          if (a.note) notes[x.email] = a.note;
        } else {
          st += att[sid] && att[sid][x.email] ? "p" : "a";
          ad += "0";
        }
      });
      return {
        id: sid,
        date: fmt(startMs, "yyyy-MM-dd"),
        time: fmt(startMs, "HH:mm"),
        st: st,
        ad: ad,
        notes: notes,
      };
    }),
  };
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
  const adminEmail = requireAdmin_();
  const report = getReport(cls, from, to);
  const sessions = report.sessions;
  const students = report.students;

  // จัดกลุ่มวันที่ตามเดือน
  const monthMap = {};
  const monthOrder = [];
  sessions.forEach(function (se, sIdx) {
    const d = new Date(se.date + "T00:00:00");
    const mKey = d.toLocaleDateString("th-TH", {
      month: "long",
      year: "numeric",
    });
    if (!monthMap[mKey]) {
      monthMap[mKey] = { name: mKey, sessions: [] };
      monthOrder.push(mKey);
    }
    monthMap[mKey].sessions.push({ session: se, index: sIdx });
  });

  // คำนวณจำนวนคอลัมน์ของตารางข้อมูลรายวัน
  // แถวข้อมูลเริ่มต้นที่แถว 3 (แถว 1=ชื่อเดือน, แถว 2=วันที่, แถว 3 เป็นต้นไป=นักเรียน)
  const startRow = 3;
  const totalStudents = students.length;
  const endRow = startRow + totalStudents - 1;

  // คอลัมน์เริ่มต้นของข้อมูลคาบเรียนคือ คอลัมน์ D (index 3 ถ้า A=0)
  // แปลง Index คอลัมน์เป็นตัวอักษร Excel (เช่น 0=A, 3=D, 26=AA)
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

  const startColIdx = 3; // คอลัมน์ D
  const sessionsCount = sessions.length;
  const endColIdx = startColIdx + sessionsCount - 1;
  const startColLet = colLetter(startColIdx);
  const endColLet = colLetter(endColIdx);

  // ตำแหน่งคอลัมน์สรุปผลที่จะอยู่ถัดจากวันสุดท้าย
  const presentColLet = colLetter(endColIdx + 1); // มาเรียน (นับตัว "/")
  const leaveColLet = colLetter(endColIdx + 2); // ป่วย/ลา (นับตัว "ล")
  const lateColLet = colLetter(endColIdx + 3); // สาย (นับตัว "ส")
  const absentColLet = colLetter(endColIdx + 4); // ขาด (นับตัว "ข")
  const notComeColLet = colLetter(endColIdx + 5); // รวมไม่มา (ลา + ขาด)
  const totalTimeColLet = colLetter(endColIdx + 6); // รวมเวลา (นับช่องที่ไม่ว่าง)
  const rateColLet = colLetter(endColIdx + 7); // ร้อยละ

  let html =
    '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">\n<head>\n<meta http-equiv="content-type" content="text/html; charset=UTF-8">\n<style>\n' +
    "table { border-collapse: collapse; font-family: sans-serif; font-size: 11pt; }\n" +
    "th, td { border: 1px solid #cbd5e1; padding: 6px 10px; text-align: center; vertical-align: middle; }\n" +
    "th { background-color: #2563eb; color: #ffffff; font-weight: bold; }\n" +
    ".name { text-align: left; }\n" +
    "</style>\n</head>\n<body>\n<table>\n";

  // แถว Header 1 (ชื่อเดือน)
  html +=
    '<tr>\n<th rowspan="2">เลขที่</th>\n<th rowspan="2">รหัส</th>\n<th rowspan="2">ชื่อ-สกุล</th>\n';
  monthOrder.forEach(function (mKey) {
    const m = monthMap[mKey];
    html +=
      '<th colspan="' + m.sessions.length + '">' + escXml_(m.name) + "</th>\n";
  });
  html +=
    '<th rowspan="2">มาเรียน</th>\n<th rowspan="2">ป่วย/ลา</th>\n<th rowspan="2">สาย</th>\n<th rowspan="2">ขาด</th>\n<th rowspan="2">รวมไม่มา</th>\n<th rowspan="2">รวมเวลา</th>\n<th rowspan="2">ร้อยละ</th>\n</tr>\n';

  // แถว Header 2 (วันที่)
  html += "<tr>\n";
  monthOrder.forEach(function (mKey) {
    const m = monthMap[mKey];
    m.sessions.forEach(function (item) {
      const d = new Date(item.session.date + "T00:00:00");
      html += "<th>" + d.getDate() + "</th>\n";
    });
  });
  html += "</tr>\n";

  // แถวข้อมูลนักเรียน (ใส่สูตร Excel ลงไปในช่องสรุปผล)
  students.forEach(function (s, idx) {
    const rNum = startRow + idx; // แถวปัจจุบันใน Excel (เริ่มแถวที่ 3)

    html += "<tr>\n";
    html += "<td>" + (idx + 1) + "</td>\n";
    html += "<td>" + escXml_(s.studentId) + "</td>\n";
    html += '<td class="name">' + escXml_(s.name) + "</td>\n";

    // ข้อมูลสถานะแต่ละคาบ
    sessions.forEach(function (se) {
      const code = se.st[idx];
      let sym = "";
      if (code === "p") sym = "/";
      else if (code === "l") sym = "ส";
      else if (code === "v") sym = "ล";
      else if (code === "a") sym = "ข";
      html += "<td>" + sym + "</td>\n";
    });

    // สูตร Excel สำหรับช่องสรุปผลท้ายตาราง
    // 1. มาเรียน (นับตัว "/"): =COUNTIF(D3:J3, "/")
    html +=
      "<td>=COUNTIF(" +
      startColLet +
      rNum +
      ":" +
      endColLet +
      rNum +
      ',"/")</td>\n';
    // 2. ป่วย/ลา (นับตัว "ล"): =COUNTIF(D3:J3, "ล")
    html +=
      "<td>=COUNTIF(" +
      startColLet +
      rNum +
      ":" +
      endColLet +
      rNum +
      ',"ล")</td>\n';
    // 3. สาย (นับตัว "ส"): =COUNTIF(D3:J3, "ส")
    html +=
      "<td>=COUNTIF(" +
      startColLet +
      rNum +
      ":" +
      endColLet +
      rNum +
      ',"ส")</td>\n';
    // 4. ขาด (นับตัว "ข"): =COUNTIF(D3:J3, "ข")
    html +=
      "<td>=COUNTIF(" +
      startColLet +
      rNum +
      ":" +
      endColLet +
      rNum +
      ',"ข")</td>\n';
    // 5. รวมไม่มา (ลา + ขาด): =L3+N3 (ปรับตามตำแหน่งคอลัมน์ใหม่)
    html +=
      "<td>=" + leaveColLet + rNum + "+" + absentColLet + rNum + "</td>\n";
    // 6. รวมเวลา: =COUNTA(D3:J3)
    html +=
      "<td>=COUNTA(" + startColLet + rNum + ":" + endColLet + rNum + ")</td>\n";
    // 7. ร้อยละ: ((มา + สาย) / รวมเวลา)
    html +=
      "<td>=IF(" +
      totalTimeColLet +
      rNum +
      ">0, (((" +
      presentColLet +
      rNum +
      "+" +
      lateColLet +
      rNum +
      "))/" +
      totalTimeColLet +
      rNum +
      ")*100, 0)</td>\n";

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
