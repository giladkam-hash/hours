(function () {
  const STORE_KEY = "hours.v1";
  const HE_MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];
  const HE_DAYS = ["א׳", "ב׳", "ג׳", "ד׳", "ה׳", "ו׳", "ש׳"];
  const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const VERSION = "4";
  const SUGGESTIONS = 6, MAX_PINS = 3, FIRST_YEAR = 2024;
  const DAY_TYPES = { eve: "ערב חג", holiday: "חג", vacation: "חופש", sick: "מחלה" };
  const STATUS = { frozen: "בהקפאה", done: "הסתיים" };
  // Rides in the projects list, so aliases, day marks and preferences reach the backup without changing Code.gs.
  const META_NUM = "__app";
  // Value of loadedAt for reports imported from old attendance files. They count in totals and never go to Excel.
  const HISTORY = "history";
  const SORTS = [["last", "דיווח אחרון"], ["name", "שם"], ["hours", "שעות"], ["num", "מספר"]];
  const FILTERS = [["active", "פעילים"], ["pinned", "נעוצים"], ["recent", "90 יום"], ["frozen", "בהקפאה"], ["done", "הסתיימו"]];
  const FONTS = [[0.92, "קטן"], [1, "רגיל"], [1.12, "גדול"], [1.25, "גדול מאוד"]];
  const MODES = [["auto", "לפי המכשיר"], ["light", "בהיר"], ["dark", "כהה"]];
  const LIGHT = { ink: "#17232f", muted: "#5f6a78", "on-accent": "#ffffff", ok: "#1f7349", "ok-soft": "#dff0e7", mark: "#a86205", "mark-soft": "#fbeed7", danger: "#b8322a", "danger-soft": "#fae6e3", rest: "#5361ad", "rest-soft": "#e4e7f6" };
  const DARK = { ink: "#e7edf3", muted: "#9dabba", "on-accent": "#0b1520", ok: "#6fcf9b", "ok-soft": "#183228", mark: "#f0b04f", "mark-soft": "#3a2c14", danger: "#f19088", "danger-soft": "#3b2120", rest: "#a4b1f0", "rest-soft": "#252b47" };
  const THEMES = {
    blueprint: { name: "כחול תוכנית",
      light: { bg: "#f2efe7", paper: "#fffdf8", line: "#ded8ca", sheet: "#14375a", accent: "#17609f", "accent-soft": "#e0ebf5" },
      dark: { bg: "#0f151c", paper: "#18212b", line: "#2c3845", sheet: "#10283f", accent: "#6cabec", "accent-soft": "#1c3047" } },
    concrete: { name: "בטון",
      light: { bg: "#e9e9e6", paper: "#fbfbf9", line: "#d3d3cd", sheet: "#3e4347", accent: "#3f5c70", "accent-soft": "#e1e7ea" },
      dark: { bg: "#131516", paper: "#1e2123", line: "#33383b", sheet: "#24282b", accent: "#a3bccb", "accent-soft": "#29323a" } },
    terracotta: { name: "טרקוטה",
      light: { bg: "#f6ece3", paper: "#fffaf5", line: "#e8d5c6", sheet: "#8a3d23", accent: "#a8482a", "accent-soft": "#f7e1d6" },
      dark: { bg: "#17100d", paper: "#241915", line: "#40302a", sheet: "#4c2415", accent: "#f0a07f", "accent-soft": "#43271c" } },
    olive: { name: "זית",
      light: { bg: "#eef0e3", paper: "#fcfdf6", line: "#d8dbc6", sheet: "#3f5233", accent: "#4c6a37", "accent-soft": "#e3ebd8" },
      dark: { bg: "#11140e", paper: "#1b2016", line: "#333a2a", sheet: "#27331f", accent: "#aecb8b", "accent-soft": "#2b3622" } },
    plum: { name: "שזיף",
      light: { bg: "#f2ecf2", paper: "#fdf9fd", line: "#e1d4e1", sheet: "#4a2a56", accent: "#74408a", "accent-soft": "#eee0f2" },
      dark: { bg: "#141017", paper: "#201a24", line: "#3a2f40", sheet: "#33203c", accent: "#cfa3e2", "accent-soft": "#35263e" } },
  };

  const $ = id => document.getElementById(id);
  const pad = n => String(n).padStart(2, "0");
  const iso = d => d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  const parse = k => new Date(k + "T00:00");
  const fmtH = h => String(Math.round(h * 100) / 100);
  const shortDate = k => { const [, m, d] = k.split("-"); return d + "/" + m; };
  const dayLabel = k => HE_DAYS[parse(k).getDay()] + " " + shortDate(k);
  const newId = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now() + "-" + Math.random().toString(16).slice(2));

  let todayIso = iso(new Date());
  let viewY = parse(todayIso).getFullYear(), viewM = parse(todayIso).getMonth();
  let sendY = viewY, sendM = viewM, historyUntil = "", filter = "active", showAll = false, lastTap = { tab: "", at: 0 };
  let selDate = todayIso, selNum = null, hours = 1, picking = true, scrolledTo = null, editingId = null;
  let unchecked = new Set(), weekExtra = new Set(), sending = false, connecting = false, currentTab = "report", lastTab = "report";

  let projects = [], entries = [], aliases = {}, dayMarks = {}, pins = [], status = {};
  let settings = { scriptUrl: "", key: "", recipient: "", employeeNum: "", lastBackup: null, backupPending: false, sort: null, font: null, theme: null, mode: null };
  let backupError = null, storageFailed = false, backupRunning = false, changes = 0, backupTimer = null;

  function toast(text) {
    const el = document.createElement("div");
    el.className = "toast"; el.setAttribute("role", "status"); el.textContent = text;
    document.body.appendChild(el);
    setTimeout(() => el.remove(), 2600);
  }

  /* Bottom sheet that resolves with the value of the button pressed. fill(body, done) can add fields or its own choices. */
  function ask(title, lines, buttons, fill) {
    return new Promise(resolve => {
      const done = value => { $("modal").hidden = true; resolve(value); };
      $("mtitle").textContent = title;
      const body = $("mbody"), box = $("mbtns");
      body.textContent = ""; box.textContent = "";
      lines.forEach(text => { const p = document.createElement("p"); p.textContent = text; body.appendChild(p); });
      if (fill) fill(body, done);
      buttons.forEach(b => {
        const el = document.createElement("button");
        el.className = (b.primary ? "primary" : "secondary") + (b.danger ? " danger" : "");
        el.textContent = b.label;
        el.onclick = () => done(b.value);
        box.appendChild(el);
      });
      $("modal").hidden = false;
    });
  }

  /* Row of buttons where one is selected. items are [value, text] pairs. */
  function choices(id, items, current, onPick) {
    const box = $(id);
    box.textContent = "";
    items.forEach(([value, text]) => {
      const b = document.createElement("button");
      b.className = "opt";
      b.textContent = text;
      b.setAttribute("aria-pressed", String(value === current));
      b.onclick = () => onPick(value);
      box.appendChild(b);
    });
  }

  // MARK: Look

  const lookMode = () => settings.mode || "auto";
  const lookTheme = () => (THEMES[settings.theme] ? settings.theme : "blueprint");
  const darkQuery = matchMedia("(prefers-color-scheme: dark)");

  function applyLook() {
    const dark = lookMode() === "dark" || (lookMode() === "auto" && darkQuery.matches);
    const colors = Object.assign({}, dark ? DARK : LIGHT, THEMES[lookTheme()][dark ? "dark" : "light"]);
    const root = document.documentElement;
    Object.entries(colors).forEach(([name, value]) => root.style.setProperty("--" + name, value));
    root.style.colorScheme = dark ? "dark" : "light";
    root.style.fontSize = (settings.font || 1) * 100 + "%";
    document.querySelector('meta[name="theme-color"]').content = colors.sheet;
  }

  function setLook(name, value) {
    settings[name] = value;
    persist();
    render();
  }

  // MARK: Holidays

  const hebrewFormat = (() => {
    try { return new Intl.DateTimeFormat("en-u-ca-hebrew", { day: "numeric", month: "long" }); } catch (e) { return null; }
  })();

  function hebrewDate(date) {
    const parts = hebrewFormat.formatToParts(date);
    return { day: Number(parts.find(p => p.type === "day").value), month: parts.find(p => p.type === "month").value };
  }

  const FIXED_DAYS = {
    Elul: { 29: "ערב ראש השנה" },
    Tishri: { 1: "ראש השנה", 2: "ראש השנה", 9: "ערב יום כיפור", 10: "יום כיפור", 14: "ערב סוכות", 15: "סוכות", 16: "חול המועד סוכות", 17: "חול המועד סוכות",
      18: "חול המועד סוכות", 19: "חול המועד סוכות", 20: "חול המועד סוכות", 21: "הושענא רבה", 22: "שמחת תורה" },
    Tevet: { 10: "צום עשרה בטבת" },
    Shevat: { 15: 'ט"ו בשבט' },
    Nisan: { 14: "ערב פסח", 15: "פסח", 16: "חול המועד פסח", 17: "חול המועד פסח", 18: "חול המועד פסח", 19: "חול המועד פסח", 20: "ערב שביעי של פסח", 21: "שביעי של פסח" },
    Iyar: { 18: 'ל"ג בעומר', 28: "יום ירושלים" },
    Sivan: { 5: "ערב שבועות", 6: "שבועות" },
  };

  /* Days that move in Israel depending on the weekday, or span two Hebrew months. */
  function movingDay(d, h, wd) {
    const { day, month } = h;
    const fast = (name, m, n) => month.startsWith(m) && ((day === n && wd !== 6) || (day === n + 1 && wd === 0)) ? name : "";
    const found = fast("צום גדליה", "Tishri", 3) || fast('צום י"ז בתמוז', "Tam", 17) || fast("תשעה באב", "Av", 9);
    if (found) return found;
    if (month === "Adar" || month === "Adar II") {
      if ((day === 13 && wd !== 6) || (day === 11 && wd === 4)) return "תענית אסתר";
      if (day === 14) return "פורים";
      if (day === 15) return "שושן פורים";
    }
    if (month === "Nisan" && ((day === 27 && wd !== 5 && wd !== 0) || (day === 26 && wd === 4) || (day === 28 && wd === 1))) return "יום השואה";
    if (month === "Iyar") {
      if ((day === 5 && (wd === 3 || wd === 4)) || (day === 4 && wd === 4) || (day === 3 && wd === 4) || (day === 6 && wd === 2)) return "יום העצמאות";
      if ((day === 4 && (wd === 2 || wd === 3)) || (day === 3 && wd === 3) || (day === 2 && wd === 3) || (day === 5 && wd === 1)) return "יום הזיכרון";
    }
    if ((month === "Kislev" && day >= 25) || (month === "Tevet" && day <= 3)) {
      for (let back = 0; back < 8; back++) {
        const earlier = new Date(d);
        earlier.setDate(earlier.getDate() - back);
        const e = hebrewDate(earlier);
        if (e.month === "Kislev" && e.day === 25) return "חנוכה";
      }
    }
    return "";
  }

  const holidayCache = new Map();

  /* Name only. Holidays never change the hours, since most of them are regular work days at the office. */
  function holiday(k) {
    if (!hebrewFormat) return "";
    if (!holidayCache.has(k)) {
      const d = parse(k), h = hebrewDate(d);
      holidayCache.set(k, (FIXED_DAYS[h.month] || {})[h.day] || movingDay(d, h, d.getDay()));
    }
    return holidayCache.get(k);
  }

  // MARK: Storage

  function loadLocal() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (!raw) return;
      const saved = JSON.parse(raw);
      projects = saved.projects || [];
      entries = saved.entries || [];
      aliases = saved.aliases || {};
      dayMarks = saved.dayMarks || {};
      pins = saved.pins || [];
      status = saved.status || {};
      settings = Object.assign(settings, saved.settings);
    } catch (e) { storageFailed = true; }
  }

  function writeLocal() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({ projects, entries, aliases, dayMarks, pins, status, settings }));
      storageFailed = false;
    } catch (e) { storageFailed = true; }
  }

  /* Saves on the phone, then backs up to the cloud a few seconds later. */
  function persist() {
    changes += 1;
    settings.backupPending = true;
    writeLocal();
    clearTimeout(backupTimer);
    backupTimer = setTimeout(runBackup, 3000);
  }

  // MARK: Cloud backup

  const connected = () => !!(settings.lastBackup && settings.scriptUrl && settings.key);

  async function callScript(url, key, action, extra) {
    const stop = new AbortController();
    const timer = setTimeout(() => stop.abort(), 60000);
    let response;
    try {
      // text/plain keeps this a simple request, which Apps Script can answer across origins.
      response = await fetch(url, {
        method: "POST", redirect: "follow", signal: stop.signal,
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(Object.assign({ key, action }, extra)),
      });
    } catch (e) {
      throw new Error("אין חיבור לסקריפט כרגע. האפליקציה תנסה שוב אוטומטית.");
    } finally {
      clearTimeout(timer);
    }
    let reply;
    try { reply = await response.json(); } catch (e) {
      throw new Error("הסקריפט לא החזיר תשובה תקינה. צריך לוודא שפורסם עם Who has access: Anyone.");
    }
    if (reply.ok) return reply;
    if (reply.error === "key") throw new Error("המפתח הסודי לא נכון.");
    throw new Error(reply.message || "הסקריפט סירב לבקשה.");
  }

  function backupProjects() {
    const { employeeNum, recipient, sort, font, theme, mode } = settings;
    const meta = { aliases, days: dayMarks, pins, status, prefs: { employeeNum, recipient, sort, font, theme, mode } };
    return projects.concat({ num: META_NUM, name: JSON.stringify(meta) });
  }
  const upload = (url, key) => callScript(url, key, "save", { projects: backupProjects(), entries });

  /* Takes aliases, day marks and preferences from a backup. What is already on this device wins. */
  function mergeMeta(text) {
    let meta;
    try { meta = JSON.parse(text); } catch (e) { return; }
    Object.entries(meta.aliases || {}).forEach(([num, alias]) => {
      if (typeof alias === "string" && alias && !aliases[num]) aliases[num] = alias;
    });
    Object.entries(meta.days || {}).forEach(([date, type]) => {
      if (/^\d{4}-\d{2}-\d{2}$/.test(date) && Object.keys(DAY_TYPES).includes(type) && !dayMarks[date]) dayMarks[date] = type;
    });
    (meta.pins || []).forEach(num => { if (pins.length < MAX_PINS && !pins.includes(num)) pins.push(num); });
    Object.entries(meta.status || {}).forEach(([num, value]) => {
      if (Object.keys(STATUS).includes(value) && !status[num]) status[num] = value;
    });
    const prefs = meta.prefs || {};
    ["employeeNum", "recipient"].forEach(name => { if (prefs[name] && !settings[name]) settings[name] = String(prefs[name]); });
    if (!settings.sort && SORTS.some(s => s[0] === prefs.sort)) settings.sort = prefs.sort;
    if (!settings.font && FONTS.some(f => f[0] === prefs.font)) settings.font = prefs.font;
    if (!settings.theme && THEMES[prefs.theme]) settings.theme = prefs.theme;
    if (!settings.mode && MODES.some(m => m[0] === prefs.mode)) settings.mode = prefs.mode;
  }

  /* Sends everything to the backup. Failures stay pending and are retried on the next open or reconnect. */
  async function runBackup() {
    if (!connected() || backupRunning || !settings.backupPending) return;
    backupRunning = true;
    try {
      // Loops so that a change made during an upload is sent too.
      while (settings.backupPending) {
        const seen = changes;
        await upload(settings.scriptUrl, settings.key);
        settings.lastBackup = new Date().toISOString();
        if (changes === seen) settings.backupPending = false;
        backupError = null;
        writeLocal();
      }
    } catch (e) {
      backupError = e.message;
    }
    backupRunning = false;
    renderStatus();
  }

  /* Merges so that neither side loses anything. Returns how many reports were added. */
  function merge(remoteProjects, remoteEntries) {
    remoteProjects.forEach(p => {
      if (p.num === META_NUM) mergeMeta(p.name);
      else if (!project(p.num)) projects.push(p);
    });
    let added = 0;
    remoteEntries.forEach(e => {
      const i = entries.findIndex(x => x.id === e.id);
      if (i < 0) { entries.push(e); added += 1; }
      else if (!entries[i].loadedAt) entries[i] = e;
    });
    return added;
  }

  async function connect() {
    const url = $("surl").value.trim(), key = $("skey").value.trim();
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) return setConnectStatus("הכתובת צריכה להתחיל ב-https://script.google.com ולהסתיים ב-exec/.", true);
    if (!key) return setConnectStatus("צריך להדביק את המפתח הסודי.", true);
    connecting = true; renderSettings();
    setConnectStatus("מתחבר…", false);
    try {
      const snapshot = await callScript(url, key, "load", {});
      const added = merge(snapshot.projects || [], snapshot.entries || []);
      await upload(url, key);
      settings.scriptUrl = url; settings.key = key;
      settings.lastBackup = new Date().toISOString();
      settings.backupPending = false;
      backupError = null;
      writeLocal();
      connecting = false;
      render();
      setConnectStatus("הגיבוי מחובר. נוספו " + added + " דיווחים מהגיבוי.", false);
    } catch (e) {
      connecting = false; renderSettings();
      setConnectStatus(e.message, true);
    }
  }

  function setConnectStatus(text, bad) {
    const el = $("sstatus");
    el.textContent = text;
    el.className = bad ? "warn" : "note";
  }

  // MARK: Data

  const project = num => projects.find(p => p.num === num);
  const nameOf = num => { const p = project(num); return p ? p.name : num; };
  /* The name shown on screen. The Excel file always keeps the original name from nameOf. */
  const label = num => aliases[num] || nameOf(num);
  const keyOf = (y, m) => y + "-" + pad(m + 1);
  const sendKey = () => keyOf(sendY, sendM);
  const onSendTab = () => currentTab === "send";
  const monthEntries = () => entries.filter(e => e.date.startsWith(keyOf(viewY, viewM)));
  const norm = s => String(s || "").replace(/["'״׳]/g, "").toLowerCase();
  const matches = (p, q) => !q || p.num.includes(q) || norm(p.name).includes(q) || norm(aliases[p.num]).includes(q);
  const hoursOn = date => entries.reduce((s, e) => (e.date === date ? s + e.hours : s), 0);
  const isHistory = e => e.loadedAt === HISTORY;
  const isPinned = num => pins.includes(num);

  /* Last report and total hours of every project. */
  function projectStats() {
    const stats = {};
    entries.forEach(e => {
      const s = stats[e.projectNum] || (stats[e.projectNum] = { last: "", hours: 0 });
      s.hours += e.hours;
      if (e.date > s.last) s.last = e.date;
    });
    return stats;
  }

  function sortProjects(list) {
    const stats = projectStats(), mode = settings.sort || "last";
    const of = p => stats[p.num] || { last: "", hours: 0 };
    const byNum = (a, b) => (Number(b.num) || 0) - (Number(a.num) || 0);
    const rules = {
      last: (a, b) => of(b).last.localeCompare(of(a).last) || byNum(a, b),
      name: (a, b) => label(a.num).localeCompare(label(b.num), "he"),
      hours: (a, b) => of(b).hours - of(a).hours || byNum(a, b),
      num: byNum,
    };
    return list.slice().sort(rules[mode] || rules.last);
  }

  /* Attendance history ends with its last month. Days up to there were never checked, so they are not judged. */
  function refreshHistoryEnd() {
    let last = "";
    entries.forEach(e => { if (isHistory(e) && e.date > last) last = e.date; });
    if (!last) { historyUntil = ""; return; }
    const d = parse(last);
    historyUntil = iso(new Date(d.getFullYear(), d.getMonth() + 1, 0));
  }

  function setDate(k) {
    selDate = k;
    viewY = parse(k).getFullYear();
    viewM = parse(k).getMonth();
  }

  /* A holiday, vacation or sick day needs no report. A holiday eve is a short day. */
  const isRest = date => !!dayMarks[date] && dayMarks[date] !== "eve";
  const isWorkday = date => parse(date).getDay() <= 4;
  const dayMarkText = date => DAY_TYPES[dayMarks[date]] || "";

  /* Office rules: Sunday to Wednesday minimum 6 and norm 7, Thursday minimum 5 and norm 6, holiday eve 5. */
  function dayRule(date) {
    if (isRest(date)) return null;
    if (dayMarks[date] === "eve") return { min: 5, norm: 5 };
    const wd = parse(date).getDay();
    if (wd <= 3) return { min: 6, norm: 7 };
    if (wd === 4) return { min: 5, norm: 6 };
    return null;
  }

  function dayState(date) {
    if (isRest(date)) return "rest";
    const rule = dayRule(date), h = hoursOn(date);
    if (!rule) return "off";
    if (date > todayIso || date <= historyUntil) return "";
    if (h === 0) return date === todayIso ? "" : "missing";
    if (h < rule.min) return "low";
    if (h < rule.norm) return "short";
    return "ok";
  }

  function monthDays(y = viewY, m = viewM) {
    const count = new Date(y, m + 1, 0).getDate(), days = [];
    for (let d = 1; d <= count; d++) days.push(iso(new Date(y, m, d)));
    return days;
  }

  /* Workdays up to today that are missing a report or below the office hours, and days off that have hours. */
  function dayIssues() {
    const list = [];
    monthDays(sendY, sendM).filter(k => k <= todayIso && k > historyUntil).forEach(k => {
      const rule = dayRule(k), h = hoursOn(k);
      if (isRest(k) && h) list.push({ date: k, text: dayLabel(k) + ": מסומן " + dayMarkText(k) + " אבל יש " + fmtH(h) + " ש׳" });
      if (!rule) return;
      if (h === 0) list.push({ date: k, text: dayLabel(k) + ": אין דיווח" });
      else if (h < rule.min) list.push({ date: k, text: dayLabel(k) + ": " + fmtH(h) + " ש׳, מתחת למינימום " + rule.min });
      else if (h < rule.norm) list.push({ date: k, text: dayLabel(k) + ": " + fmtH(h) + " ש׳, מתחת ל-" + rule.norm + " של המשרד" });
    });
    return list;
  }

  /* Pinned projects first, then everything from the last day reported, then the most recent projects fill the rest. */
  function suggested(date) {
    const sorted = entries.filter(e => !status[e.projectNum])
      .sort((a, b) => b.date.localeCompare(a.date) || String(b.created).localeCompare(String(a.created)));
    const earlier = sorted.filter(e => e.date < date);
    const lastDay = earlier.length ? earlier[0].date : null;
    const nums = pins.slice();
    earlier.filter(e => e.date === lastDay).forEach(e => { if (!nums.includes(e.projectNum)) nums.push(e.projectNum); });
    earlier.concat(sorted).forEach(e => {
      if (nums.length < SUGGESTIONS && !nums.includes(e.projectNum)) nums.push(e.projectNum);
    });
    return nums.map(project).filter(Boolean);
  }

  /* Rows that can still go to Master Plan, by month. History is never one of them. */
  function pendingByMonth() {
    const seen = new Set(), counts = {};
    entries.filter(e => !e.loadedAt).forEach(e => {
      const id = e.date + "|" + e.projectNum, key = e.date.slice(0, 7);
      if (seen.has(id)) return;
      seen.add(id);
      counts[key] = (counts[key] || 0) + 1;
    });
    return counts;
  }

  function monthRows() {
    const by = new Map();
    entries.filter(e => e.date.startsWith(sendKey()) && !isHistory(e)).forEach(e => {
      const loaded = !!e.loadedAt, id = e.date + "|" + e.projectNum + "|" + (loaded ? 1 : 0);
      if (!by.has(id)) by.set(id, { id, date: e.date, num: e.projectNum, name: nameOf(e.projectNum), label: label(e.projectNum), hours: 0, entryIds: [], loaded });
      const row = by.get(id);
      row.hours += e.hours;
      row.entryIds.push(e.id);
    });
    return [...by.values()].sort((a, b) => a.date.localeCompare(b.date) || a.num.localeCompare(b.num));
  }

  function exportState() {
    const rows = monthRows(), pending = rows.filter(r => !r.loaded);
    return { pending, done: rows.filter(r => r.loaded), selected: pending.filter(r => !unchecked.has(r.id)) };
  }

  function setLoaded(rows, value) {
    const ids = new Set(rows.flatMap(r => r.entryIds));
    entries.forEach(e => {
      if (!ids.has(e.id)) return;
      if (value) e.loadedAt = value; else delete e.loadedAt;
    });
    persist();
    render();
  }

  /* A report that was already sent changes here only, so the user must agree before touching it. */
  async function allowSentChange(entry) {
    if (!entry.loadedAt) return true;
    return ask("הדיווח הזה כבר נשלח", ["שינוי כאן לא משנה את מה שכבר נקלט במאסטר פלן. אם הקובץ נקלט, צריך לתקן גם שם ידנית."],
      [{ label: "ביטול", value: false, primary: true }, { label: "לשנות בכל זאת", value: true }]);
  }

  /* Marks one day, or every Sunday to Thursday up to a chosen date. A holiday eve is always a single day. */
  async function markDay(date) {
    if (document.activeElement) document.activeElement.blur();
    const now = dayMarks[date] || "";
    const options = [["", "יום רגיל"], ["eve", "ערב חג (5 שעות)"], ["holiday", "חג"], ["vacation", "יום חופש"], ["sick", "יום מחלה"]];
    let until = null;
    const choice = await ask("סימון " + dayLabel(date), ["בחג, בחופש ובמחלה לא נדרש דיווח."],
      options.map(([value, text]) => ({ label: text, value, primary: value === now })).concat({ label: "ביטול", value: null }),
      body => {
        const field = document.createElement("label");
        field.className = "stack";
        field.innerHTML = '<span>עד תאריך, לכמה ימים ברצף</span><input type="date">';
        until = field.children[1];
        until.min = date; until.value = date;
        body.appendChild(field);
      });
    if (choice === null) return;
    const last = choice === "eve" || !until.value || until.value < date ? date : until.value, days = [];
    for (const d = parse(date); iso(d) <= last && days.length < 100; d.setDate(d.getDate() + 1)) {
      if (isWorkday(iso(d))) days.push(iso(d));
    }
    days.forEach(k => { if (choice) dayMarks[k] = choice; else delete dayMarks[k]; });
    persist();
    render();
    if (days.length > 1) toast("עודכנו " + days.length + " ימים");
  }

  function removeProject(num) {
    projects = projects.filter(x => x.num !== num);
    pins = pins.filter(n => n !== num);
    delete aliases[num];
    delete status[num];
    if (selNum === num) { selNum = null; picking = true; }
  }

  async function editProject(p) {
    const used = entries.some(e => e.projectNum === p.num);
    const lines = [p.name + " · " + p.num, "הכינוי מוצג רק באפליקציה. בקובץ למאסטר פלן נשאר השם המקורי."];
    if (used) lines.push("אי אפשר למחוק פרויקט עם דיווחים. אפשר לסמן אותו כהסתיים.");
    const buttons = [{ label: "שמירה", value: "save", primary: true }, { label: "ביטול", value: null }];
    if (!used) buttons.push({ label: "מחיקת פרויקט", value: "delete", danger: true });
    let alias, pin, state;
    const choice = await ask(label(p.num), lines, buttons, body => {
      const box = document.createElement("div");
      box.className = "fields";
      box.innerHTML = '<label class="stack"><span>כינוי</span><input type="text" maxlength="40" autocomplete="off" placeholder="כינוי קצר"></label>'
        + '<label class="check"><input type="checkbox"><span>נעוץ, תמיד ראשון בהמלצות</span></label>'
        + '<label class="stack"><span>מצב הפרויקט</span><select><option value="">פעיל</option><option value="frozen">בהקפאה</option><option value="done">הסתיים</option></select></label>';
      [alias, pin, state] = box.querySelectorAll("input[type=text],input[type=checkbox],select");
      alias.value = aliases[p.num] || "";
      pin.checked = isPinned(p.num);
      state.value = status[p.num] || "";
      body.appendChild(box);
    });
    if (choice === "delete") {
      const sure = await ask("למחוק את הפרויקט?", [p.name], [{ label: "ביטול", value: false, primary: true }, { label: "מחיקה", value: true, danger: true }]);
      if (!sure) return;
      removeProject(p.num);
    }
    if (choice === "save") {
      const text = alias.value.trim();
      const taken = text && projects.find(x => x.num !== p.num && aliases[x.num] === text);
      if (taken) { toast("הכינוי הזה כבר שייך לפרויקט " + taken.num); return; }
      if (text) aliases[p.num] = text; else delete aliases[p.num];
      if (state.value) status[p.num] = state.value; else delete status[p.num];
      const wantPin = pin.checked && !state.value;
      if (wantPin && !isPinned(p.num) && pins.length >= MAX_PINS) toast("אפשר לנעוץ עד " + MAX_PINS + " פרויקטים");
      else if (wantPin && !isPinned(p.num)) pins.push(p.num);
      else if (!wantPin) pins = pins.filter(n => n !== p.num);
    }
    if (choice) { persist(); render(); }
  }

  function saveProject() {
    const num = $("nnum").value.trim(), name = $("nname").value.trim(), err = $("nerr");
    const fail = text => { err.textContent = text; err.hidden = false; };
    err.hidden = true;
    if (!num || !name) return fail("צריך למלא מספר פרויקט וגם שם פרויקט.");
    if (!/^[0-9]+$/.test(num)) return fail("מספר פרויקט יכול לכלול ספרות בלבד.");
    const same = project(num);
    if (same) return fail("מספר " + num + " כבר קיים: " + same.name + ".");
    if (projects.some(p => p.name === name)) return fail("כבר יש פרויקט בשם הזה.");
    projects.push({ num, name });
    persist();
    $("nnum").value = ""; $("nname").value = "";
    selNum = num; picking = false;
    toast("הפרויקט נשמר");
    showTab("report");
  }

  // MARK: Rendering helpers

  function paragraph(text, className) {
    const el = document.createElement("p");
    el.className = className; el.textContent = text;
    return el;
  }

  function projectButton(p, className, onPick) {
    const b = document.createElement("button");
    b.className = className;
    b.innerHTML = '<span class="pname"></span><span class="pnum"></span>';
    b.children[0].textContent = label(p.num);
    b.children[1].textContent = [p.num, isPinned(p.num) ? "נעוץ" : "", STATUS[status[p.num]] || ""].filter(Boolean).join(" · ");
    b.onclick = onPick;
    return b;
  }

  // MARK: Header and days

  function renderHero() {
    const y = onSendTab() ? sendY : viewY, m = onSendTab() ? sendM : viewM, key = keyOf(y, m);
    $("mname").textContent = HE_MONTHS[m] + " " + y;
    const total = entries.reduce((s, e) => (e.date.startsWith(key) ? s + e.hours : s), 0);
    const judged = monthDays(y, m).filter(k => k <= todayIso && k > historyUntil);
    const need = judged.reduce((s, k) => { const r = dayRule(k); return s + (r ? r.norm : 0); }, 0);
    $("mtotal").textContent = fmtH(total);
    $("mneed").textContent = judged.length ? fmtH(need) : "–";
    $("mgaplabel").textContent = judged.length && total >= need ? "עודף" : "חסר";
    $("mgap").textContent = judged.length ? fmtH(Math.abs(need - total)) : "–";
    $("prog").hidden = !need;
    $("mbar").style.width = (need ? Math.min(100, total / need * 100) : 0) + "%";
  }

  function renderDays() {
    const box = $("days");
    box.textContent = "";
    monthDays().forEach(k => {
      const h = hoursOn(k), wd = parse(k).getDay();
      const b = document.createElement("button");
      b.className = ["day", dayState(k), k === todayIso ? "today" : "", k === selDate ? "sel" : "", holiday(k) ? "hol" : ""].join(" ");
      if (k === selDate) b.setAttribute("aria-current", "date");
      b.innerHTML = '<small></small><span class="n"></span><span class="h"></span>';
      b.children[0].textContent = HE_DAYS[wd];
      b.children[1].textContent = parse(k).getDate();
      b.children[2].textContent = h ? fmtH(h) : dayMarkText(k);
      b.onclick = () => { selDate = k; editingId = null; render(); };
      box.appendChild(b);
      if (k === selDate && scrolledTo !== k && currentTab === "report") {
        scrolledTo = k;
        requestAnimationFrame(() => b.scrollIntoView({ inline: "center", block: "nearest" }));
      }
    });
  }

  function renderDayBar() {
    const type = dayMarks[selDate], name = holiday(selDate), work = isWorkday(selDate);
    $("daybar").hidden = !work && !name;
    $("dhol").textContent = name;
    $("dmark").textContent = !type ? "" : type === "eve" ? "ערב חג: נדרשות 5 שעות" : dayMarkText(selDate) + ": לא נדרש דיווח";
    $("dmarkbtn").hidden = !work;
    $("dmarkbtn").textContent = type ? "שינוי" : "סימון חופש, מחלה או חג";
  }

  // MARK: Daily report

  function renderPicker() {
    const cur = project(selNum), show = !!cur && !picking;
    $("chosen").hidden = !show;
    $("picker").hidden = show;
    if (show) { $("cnum").textContent = cur.num; $("cname").textContent = label(cur.num); return; }
    const choose = p => () => { selNum = p.num; picking = false; showAll = false; $("psearch").value = ""; render(); };
    const raw = $("psearch").value.trim(), q = norm(raw), box = $("plist"), sug = $("sug");
    const recent = suggested(selDate);
    sug.textContent = "";
    sug.hidden = !!q || !recent.length;
    recent.forEach(p => sug.appendChild(projectButton(p, "sugbtn", choose(p))));
    box.textContent = "";
    $("pshow").hidden = !!q || !projects.length;
    $("pshow").textContent = showAll ? "הסתרת הרשימה" : "כל הפרויקטים";
    if (q) {
      const hits = sortProjects(projects.filter(p => matches(p, q)));
      hits.forEach(p => box.appendChild(projectButton(p, "pick", choose(p))));
      if (!hits.length) {
        box.appendChild(paragraph("לא נמצא פרויקט כזה.", "empty"));
        const add = document.createElement("button");
        add.className = "secondary"; add.textContent = "הוספת פרויקט חדש";
        add.onclick = () => { $(/^\d+$/.test(raw) ? "nnum" : "nname").value = raw; $("psearch").value = ""; showTab("projects"); };
        box.appendChild(add);
      }
      return;
    }
    if (!projects.length) { box.appendChild(paragraph("עדיין אין פרויקטים. הרשימה תיטען אחרי חיבור הגיבוי.", "empty")); return; }
    if (!showAll) return;
    sortProjects(projects.filter(p => !recent.includes(p) && !status[p.num])).forEach(p => box.appendChild(projectButton(p, "pick", choose(p))));
  }

  function renderHours() {
    $("hours").textContent = fmtH(hours);
    const quick = $("quick");
    quick.textContent = "";
    [1, 2, 3, 4, 5, 6, 7].forEach(v => {
      const b = document.createElement("button");
      b.className = "chip" + (v === hours ? " on" : "");
      b.textContent = v;
      b.onclick = () => { hours = v; renderHours(); };
      quick.appendChild(b);
    });
    $("add").textContent = "הוספה ליום " + dayLabel(selDate);
    $("add").disabled = !project(selNum) || picking;
  }

  function renderLog() {
    $("dayh").textContent = "דיווחים ליום " + dayLabel(selDate);
    const list = entries.filter(e => e.date === selDate), box = $("log");
    box.textContent = "";
    const rule = dayRule(selDate), total = list.reduce((s, e) => s + e.hours, 0);
    $("dtotal").textContent = list.length ? fmtH(total) + (rule ? " / " + rule.norm : "") + " ש׳" : "";
    if (!list.length) { box.appendChild(paragraph("אין דיווחים ליום הזה.", "empty")); return; }
    list.forEach(e => {
      const row = document.createElement("div");
      row.className = "item";
      row.innerHTML = '<span class="p"></span><button class="hbtn"></button><button class="del" aria-label="מחיקת הדיווח">×</button>';
      row.children[0].textContent = label(e.projectNum);
      if (e.loadedAt) {
        const mark = document.createElement("span");
        mark.className = "sentmark"; mark.textContent = isHistory(e) ? "היסטוריה" : "נשלח";
        row.insertBefore(mark, row.children[1]);
      }
      const hbtn = row.querySelector(".hbtn");
      hbtn.textContent = fmtH(e.hours);
      hbtn.setAttribute("aria-label", "שינוי השעות של " + label(e.projectNum));
      hbtn.onclick = async () => {
        if (editingId === e.id) { editingId = null; render(); return; }
        if (await allowSentChange(e)) { editingId = e.id; renderLog(); }
      };
      row.querySelector(".del").onclick = async () => {
        if (!await allowSentChange(e)) return;
        entries = entries.filter(x => x.id !== e.id);
        persist();
        render();
      };
      if (editingId === e.id) row.appendChild(editor(e));
      box.appendChild(row);
    });
  }

  function editor(e) {
    const box = document.createElement("div");
    box.className = "edit";
    box.innerHTML = '<button class="step" aria-label="פחות חצי שעה">−</button><output></output><button class="step" aria-label="עוד חצי שעה">+</button><button class="linkbtn">סיום</button>';
    box.children[1].textContent = fmtH(e.hours);
    const change = step => {
      e.hours = Math.min(23.5, Math.max(0.5, e.hours + step));
      persist();
      render();
    };
    box.children[0].onclick = () => change(-0.5);
    box.children[2].onclick = () => change(0.5);
    box.children[3].onclick = () => { editingId = null; render(); };
    return box;
  }

  // MARK: Weekly report

  function weekDates() {
    const start = parse(selDate);
    start.setDate(start.getDate() - start.getDay());
    return [0, 1, 2, 3, 4].map(i => { const d = new Date(start); d.setDate(d.getDate() + i); return iso(d); });
  }

  function moveWeek(step) {
    const d = parse(weekDates()[0]);
    d.setDate(d.getDate() + step * 7);
    setDate(iso(d));
    weekExtra = new Set();
    render();
  }

  /* One cell holds one project on one day. Sent reports are never touched from here. */
  function setCell(date, num, value) {
    const open = entries.filter(e => e.date === date && e.projectNum === num && !e.loadedAt);
    const keep = open[0];
    entries = entries.filter(e => !open.includes(e) || (e === keep && value > 0));
    if (value > 0) {
      if (keep) keep.hours = value;
      else entries.push({ id: newId(), date, projectNum: num, hours: value, created: new Date().toISOString() });
    }
    persist();
  }

  function renderWeekTotals(dates) {
    const box = $("wtotals");
    box.textContent = "";
    dates.forEach(k => {
      const el = document.createElement("div");
      el.className = "wtotal " + dayState(k);
      el.textContent = fmtH(hoursOn(k));
      box.appendChild(el);
    });
  }

  function weekCell(date, num) {
    const input = document.createElement("input");
    const list = entries.filter(e => e.date === date && e.projectNum === num);
    const sent = list.some(e => e.loadedAt), total = list.reduce((s, e) => s + e.hours, 0);
    input.type = "text"; input.inputMode = "decimal"; input.autocomplete = "off";
    input.value = total ? fmtH(total) : "";
    input.disabled = sent;
    input.classList.toggle("rest", isRest(date));
    input.setAttribute("aria-label", label(num) + " " + dayLabel(date));
    input.onchange = () => {
      const raw = input.value.trim().replace(",", "."), value = raw === "" ? 0 : Math.round(Number(raw) * 2) / 2;
      if (!(value >= 0 && value < 24)) { toast("שעות צריכות להיות מספר מתחת ל-24"); input.value = total ? fmtH(total) : ""; return; }
      input.value = value ? fmtH(value) : "";
      setCell(date, num, value);
      renderWeekTotals(weekDates());
      renderHero();
    };
    return input;
  }

  function renderWeek() {
    const dates = weekDates();
    $("wname").textContent = shortDate(dates[0]) + " – " + shortDate(dates[4]);
    const head = $("whead");
    head.textContent = "";
    dates.forEach(k => {
      const el = document.createElement("button");
      el.className = "wday" + (dayMarks[k] ? " marked" : "") + (holiday(k) ? " hol" : "");
      el.innerHTML = "<b></b><small></small>";
      el.children[0].textContent = HE_DAYS[parse(k).getDay()];
      el.children[1].textContent = dayMarkText(k) || shortDate(k);
      el.setAttribute("aria-label", "סימון " + dayLabel(k) + (holiday(k) ? ", " + holiday(k) : ""));
      el.onclick = () => markDay(k);
      head.appendChild(el);
    });
    const nums = [];
    entries.filter(e => dates.includes(e.date)).forEach(e => { if (!nums.includes(e.projectNum)) nums.push(e.projectNum); });
    weekExtra.forEach(n => { if (!nums.includes(n)) nums.push(n); });
    suggested(dates[0]).forEach(p => { if (!nums.includes(p.num)) nums.push(p.num); });
    const rows = $("wrows");
    rows.textContent = "";
    if (!nums.length) rows.appendChild(paragraph("מוסיפים פרויקט למטה כדי להתחיל.", "empty"));
    nums.forEach(num => {
      const row = document.createElement("div");
      row.className = "wproject";
      row.innerHTML = '<div class="row"><span class="pname"></span><span class="pnum"></span></div><div class="wgrid"></div>';
      row.querySelector(".pname").textContent = label(num);
      row.querySelector(".pnum").textContent = num;
      dates.forEach(k => row.querySelector(".wgrid").appendChild(weekCell(k, num)));
      rows.appendChild(row);
    });
    renderWeekTotals(dates);
    renderWeekSearch();
  }

  function renderWeekSearch() {
    const q = norm($("wsearch").value.trim()), box = $("wlist");
    box.textContent = "";
    if (!q) return;
    sortProjects(projects.filter(p => matches(p, q))).slice(0, 8).forEach(p => {
      box.appendChild(projectButton(p, "pick", () => { weekExtra.add(p.num); $("wsearch").value = ""; renderWeek(); }));
    });
  }

  // MARK: Month

  function renderCalendar() {
    const box = $("cal");
    box.textContent = "";
    HE_DAYS.forEach(d => { const el = document.createElement("div"); el.className = "head"; el.textContent = d; box.appendChild(el); });
    const days = monthDays();
    for (let i = 0; i < parse(days[0]).getDay(); i++) box.appendChild(document.createElement("div"));
    days.forEach(k => {
      const h = hoursOn(k), b = document.createElement("button");
      b.className = ["cell", dayState(k), k === todayIso ? "today" : "", holiday(k) ? "hol" : ""].join(" ");
      b.innerHTML = '<span class="n"></span><span class="h"></span>';
      b.children[0].textContent = parse(k).getDate();
      b.children[1].textContent = h ? fmtH(h) : dayMarkText(k);
      b.children[1].classList.toggle("t", !h);
      b.setAttribute("aria-label", dayLabel(k) + (holiday(k) ? ", " + holiday(k) : "") + (dayMarks[k] ? ", " + dayMarkText(k) : "") + (h ? ", " + fmtH(h) + " שעות" : ""));
      b.onclick = () => { selDate = k; scrolledTo = null; showTab("report"); };
      box.appendChild(b);
    });
    const names = [], list = $("holist");
    days.forEach(k => { const name = holiday(k); if (name && !names.some(n => n.name === name)) names.push({ name, date: k }); });
    list.textContent = names.map(n => shortDate(n.date) + " " + n.name).join(" · ");
    list.hidden = !names.length;
  }

  function renderSum() {
    const by = {}, box = $("sum");
    monthEntries().forEach(e => { by[e.projectNum] = (by[e.projectNum] || 0) + e.hours; });
    const rows = Object.entries(by).sort((a, b) => b[1] - a[1]);
    box.textContent = "";
    if (!rows.length) { box.appendChild(paragraph("אין עדיין שעות בחודש הזה.", "empty")); return; }
    const table = document.createElement("table");
    table.className = "sum";
    rows.forEach(([num, h]) => {
      const tr = document.createElement("tr");
      tr.innerHTML = "<td></td><td></td>";
      tr.children[0].textContent = label(num);
      tr.children[1].textContent = fmtH(h);
      table.appendChild(tr);
    });
    box.appendChild(table);
  }

  function exportProblems(selected) {
    const list = [];
    if (!settings.employeeNum) list.push("חסר מספר עובד. ממלאים אותו בהגדרות.");
    selected.filter(r => r.hours >= 24).forEach(r => {
      list.push("ב-" + shortDate(r.date) + " רשומות " + fmtH(r.hours) + " שעות על " + r.label + ". הקובץ תומך בפחות מ-24 שעות בשורה.");
    });
    return list;
  }

  function renderExport() {
    const { pending, done, selected } = exportState(), box = $("xrows");
    const current = todayIso.slice(0, 7), waiting = pendingByMonth();
    const older = Object.keys(waiting).filter(k => k < current).sort();
    $("xtitle").textContent = "דיווח חודשי: " + HE_MONTHS[sendM] + " " + sendY;
    $("xnow").hidden = sendKey() === current;
    $("senddot").hidden = !older.length;
    const old = $("xold");
    old.textContent = "";
    older.filter(k => k !== sendKey()).forEach(k => {
      const [y, m] = k.split("-").map(Number), b = document.createElement("button");
      b.className = "reminder";
      b.textContent = "ב" + HE_MONTHS[m - 1] + " " + y + " נשארו שורות שלא נשלחו: " + waiting[k];
      b.onclick = () => goMonth(y, m - 1);
      old.appendChild(b);
    });
    box.textContent = "";
    if (!pending.length && !done.length) box.appendChild(paragraph("אין דיווחים לשליחה בחודש הזה.", "empty"));
    pending.forEach(r => {
      const row = document.createElement("label");
      row.className = "xrow";
      row.innerHTML = '<input type="checkbox"><span class="d"></span><span class="p"></span><span class="hh"></span>';
      row.children[0].checked = !unchecked.has(r.id);
      row.children[0].onchange = e => { if (e.target.checked) unchecked.delete(r.id); else unchecked.add(r.id); renderExport(); };
      row.children[1].textContent = shortDate(r.date);
      row.children[2].textContent = r.label;
      row.children[3].textContent = fmtH(r.hours);
      box.appendChild(row);
    });
    done.forEach(r => {
      const row = document.createElement("div");
      row.className = "xrow done";
      row.innerHTML = '<span class="sentmark">נשלח</span><span class="d"></span><span class="p"></span><span class="hh"></span><button class="linkbtn">החזרה</button>';
      row.children[1].textContent = shortDate(r.date);
      row.children[2].textContent = r.label;
      row.children[3].textContent = fmtH(r.hours);
      row.children[4].setAttribute("aria-label", "החזרת " + r.label + " " + shortDate(r.date) + " לרשימה");
      row.children[4].onclick = async () => {
        const back = await ask("להחזיר לשליחה?", ["השורה כבר נשלחה. אם הקובץ נקלט במאסטר פלן, שליחה נוספת עלולה להכפיל את השעות."],
          [{ label: "ביטול", value: false, primary: true }, { label: "להחזיר לרשימה", value: true }]);
        if (back) setLoaded([r], null);
      };
      box.appendChild(row);
    });

    const problems = exportProblems(selected), messages = problems.slice();
    if (!settings.recipient) messages.push("חסרה כתובת מייל. ממלאים אותה בהגדרות.");
    selected.filter(r => done.some(d => d.date === r.date && d.num === r.num)).forEach(r => {
      messages.push("ב-" + shortDate(r.date) + " כבר נשלחו שעות על " + r.label + ". השורה הזאת היא תוספת בלבד.");
    });
    const byDay = {};
    entries.filter(e => e.date.startsWith(sendKey())).forEach(e => { byDay[e.date] = (byDay[e.date] || 0) + e.hours; });
    Object.keys(byDay).sort().forEach(k => {
      if (byDay[k] > 12) messages.push("ב-" + shortDate(k) + " רשומות " + fmtH(byDay[k]) + " שעות. כדאי לוודא שאין דיווח כפול.");
    });
    const warn = $("xwarn");
    warn.textContent = "";
    messages.forEach(text => warn.appendChild(paragraph(text, "warn")));

    const canExport = selected.length > 0 && !problems.length && !sending;
    const count = selected.length === 1 ? "שורה אחת" : selected.length + " שורות";
    $("xsend").textContent = sending ? "שולח…" : "שליחה במייל (" + count + ")";
    $("xsend").disabled = !canExport || !settings.recipient || !connected();
    $("xshare").disabled = !canExport;
    $("xdone").textContent = "סימון " + count + " כנשלחו";
    $("xdone").disabled = !selected.length;
    $("xall").textContent = selected.length ? "ניקוי הבחירה" : "בחירת הכל";
    $("xall").hidden = !pending.length;
    $("xnote").textContent = (settings.recipient ? "המייל נשלח אל " + settings.recipient + ". " : "")
      + "שורות שנשלחו מסומנות אוטומטית ולא נכנסות לקובץ הבא. במאסטר פלן: שעות לפרויקטים, תצוגת רשימה, קליטת שעות מקובץ אקסל.";
  }

  // MARK: Projects and settings

  function renderProjectList() {
    const q = norm($("asearch").value.trim()), box = $("pall"), stats = projectStats();
    const since = new Date(); since.setDate(since.getDate() - 90);
    const rules = {
      active: p => !status[p.num],
      pinned: p => isPinned(p.num),
      recent: p => !!stats[p.num] && stats[p.num].last >= iso(since),
      frozen: p => status[p.num] === "frozen",
      done: p => status[p.num] === "done",
    };
    choices("psort", SORTS, settings.sort || "last", value => setLook("sort", value));
    choices("pfilter", FILTERS, filter, value => { filter = value; renderProjectList(); });
    box.textContent = "";
    const list = sortProjects(projects.filter(p => (q ? matches(p, q) : rules[filter](p))));
    $("pcount").textContent = list.length === 1 ? "פרויקט אחד" : list.length + " פרויקטים";
    if (!list.length) { box.appendChild(paragraph("לא נמצאו פרויקטים.", "empty")); return; }
    list.forEach(p => {
      const s = stats[p.num], row = document.createElement("div");
      row.className = "prow";
      row.innerHTML = '<span class="pname"><span></span><small class="porig"></small><small class="porig"></small></span><span class="pnum"></span>'
        + '<button class="pencil"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 20l1-4L16.5 4.5a2.1 2.1 0 0 1 3 3L8 19z"/><path d="M14.5 6.5l3 3"/></svg></button>';
      const small = row.querySelectorAll(".porig");
      row.querySelector(".pname span").textContent = label(p.num);
      small[0].textContent = aliases[p.num] ? p.name : "";
      small[1].textContent = [isPinned(p.num) ? "נעוץ" : "", STATUS[status[p.num]] || "", s ? fmtH(s.hours) + " ש׳" : "", s ? "אחרון " + shortDate(s.last) + "/" + s.last.slice(2, 4) : ""].filter(Boolean).join(" · ");
      row.querySelector(".pnum").textContent = p.num;
      const edit = row.querySelector(".pencil");
      edit.setAttribute("aria-label", "עריכת " + p.name);
      edit.onclick = () => editProject(p);
      box.appendChild(row);
    });
  }

  function renderSettings() {
    const standalone = window.navigator.standalone === true || matchMedia("(display-mode: standalone)").matches;
    $("installcard").hidden = standalone;
    for (const [id, value] of [["surl", settings.scriptUrl], ["skey", settings.key], ["semp", settings.employeeNum], ["smail", settings.recipient]]) {
      if (document.activeElement !== $(id) && !$(id).dataset.dirty) $(id).value = value;
    }
    $("sconnect").disabled = connecting;
    $("stest").disabled = sending || !connected() || !settings.recipient;
    if (!connected()) $("cloud").open = true;
    choices("sfont", FONTS.map(([value, text]) => [value, text]), settings.font || 1, value => setLook("font", value));
    choices("smode", MODES, lookMode(), value => setLook("mode", value));
    choices("stheme", Object.keys(THEMES).map(id => [id, THEMES[id].name]), lookTheme(), value => setLook("theme", value));
    [...$("stheme").children].forEach((b, i) => { b.style.setProperty("--swatch", THEMES[Object.keys(THEMES)[i]].light.sheet); });
    $("sversion").textContent = "גרסה " + VERSION;
  }

  function renderStatus() {
    const banner = $("banner");
    let text = "";
    if (storageFailed) text = "הדפדפן לא מצליח לשמור נתונים במכשיר. צריך לבדוק את הגדרות הפרטיות של Safari.";
    else if (!connected()) text = "הגיבוי לא מחובר. בלי חיבור הנתונים נשמרים רק במכשיר הזה, ורשימת הפרויקטים ריקה. מחברים אותו בהגדרות.";
    else if (backupError) text = backupError;
    banner.textContent = text;
    banner.hidden = !text;
    if (connected() && !connecting && !$("sstatus").classList.contains("warn")) {
      const when = new Date(settings.lastBackup);
      setConnectStatus((settings.backupPending ? "יש שינויים שממתינים לגיבוי. " : "") + "גיבוי אחרון: " + shortDate(iso(when)) + " " + pad(when.getHours()) + ":" + pad(when.getMinutes()), false);
    }
  }

  function render() {
    refreshHistoryEnd(); applyLook();
    renderHero(); renderDays(); renderDayBar(); renderPicker(); renderHours(); renderLog();
    if (document.activeElement === null || !document.activeElement.closest("#wrows")) renderWeek();
    renderCalendar(); renderSum(); renderExport(); renderProjectList(); renderSettings(); renderStatus();
  }

  function showTab(name) {
    if (name !== "settings") lastTab = name;
    currentTab = name;
    ["report", "week", "month", "send", "projects", "settings"].forEach(t => { $("tab-" + t).hidden = t !== name; });
    document.querySelectorAll(".tabbar button").forEach(b => {
      if (b.dataset.tab === name) b.setAttribute("aria-current", "page"); else b.removeAttribute("aria-current");
    });
    $("monthbar").hidden = !["report", "month", "send"].includes(name);
    editingId = null;
    render();
    window.scrollTo(0, 0);
  }

  // MARK: Export actions

  /* Warns about missing or short days. Returns false when the user chose to fix them first. */
  async function reviewDays() {
    const issues = dayIssues();
    if (!issues.length) return true;
    const choice = await ask("נמצאו " + issues.length + " ימים לבדיקה", issues.map(i => i.text), [
      { label: "לתקן את הימים", value: "fix", primary: true },
      { label: "לשלוח בכל זאת", value: "send" },
    ]);
    if (choice === "send") return true;
    setDate(issues[0].date); scrolledTo = null;
    showTab("report");
    return false;
  }

  function exportFile(selected) {
    return { name: "master-plan-hours-" + sendKey() + ".xlsx", bytes: HoursXlsx.file(settings.employeeNum, selected) };
  }

  function base64(bytes) {
    let text = "";
    bytes.forEach(b => { text += String.fromCharCode(b); });
    return btoa(text);
  }

  async function sendMail() {
    if (!await reviewDays()) return;
    const { selected } = exportState();
    const title = HE_MONTHS[sendM] + " " + sendY, file = exportFile(selected);
    sending = true; renderExport();
    try {
      await callScript(settings.scriptUrl, settings.key, "mail", {
        to: settings.recipient, subject: "דיווח שעות למאסטר פלן - " + title,
        text: "מצורף קובץ שעות ל" + title + ", " + selected.length + " שורות.",
        filename: file.name, file: base64(file.bytes),
      });
      sending = false;
      setLoaded(selected, new Date().toISOString());
      toast("המייל נשלח אל " + settings.recipient);
    } catch (e) {
      sending = false;
      render();
      toast(e.message);
    }
  }

  async function testMail() {
    sending = true; renderSettings();
    try {
      await callScript(settings.scriptUrl, settings.key, "mail", {
        to: settings.recipient, subject: "בדיקה מיומן שעות", text: "אם המייל הזה הגיע, שליחת הקובץ תעבוד.",
      });
      toast("מייל הבדיקה נשלח");
    } catch (e) {
      toast(e.message);
    }
    sending = false; render();
  }

  /* Returns true when the file left the app through the share sheet or a download. */
  async function saveFile(file) {
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file] }); return true; } catch (e) {
        if (e.name !== "AbortError") toast("השיתוף נכשל");
        return false;
      }
    }
    const link = document.createElement("a");
    link.href = URL.createObjectURL(file);
    link.download = file.name;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 10000);
    return true;
  }

  async function shareFile() {
    if (!await reviewDays()) return;
    const { selected } = exportState(), data = exportFile(selected);
    if (!await saveFile(new File([data.bytes], data.name, { type: XLSX_TYPE }))) return;
    const mark = await ask("לסמן את השורות כנשלחו?", ["כך הן לא ייכנסו לקובץ הבא."],
      [{ label: "כן, לסמן " + selected.length + " שורות", value: true, primary: true }, { label: "לא עכשיו", value: false }]);
    if (mark) setLoaded(selected, new Date().toISOString());
  }

  function saveBackupCopy() {
    const text = JSON.stringify({ projects: backupProjects(), entries });
    saveFile(new File([text], "hours-backup-" + todayIso + ".json", { type: "application/json" }));
  }

  async function importBackup(file) {
    try {
      const data = JSON.parse(await file.text());
      if (!Array.isArray(data.projects) || !Array.isArray(data.entries)) throw new Error("shape");
      const added = merge(data.projects.filter(p => p && p.num && p.name),
        data.entries.filter(e => e && e.id && e.date && e.projectNum && typeof e.hours === "number"));
      persist();
      render();
      toast("הייבוא הסתיים. נוספו " + added + " דיווחים.");
    } catch (e) {
      toast("הקובץ הזה לא נראה כמו קובץ גיבוי.");
    }
  }

  // MARK: Events

  /* The send tab keeps its own month, so choosing what to send never moves the daily report. */
  function goMonth(y, m) {
    if (onSendTab()) { sendY = y; sendM = m; render(); return; }
    setDate(keyOf(y, m) === todayIso.slice(0, 7) ? todayIso : iso(new Date(y, m, 1)));
    scrolledTo = null;
    render();
  }

  function moveMonth(step) {
    const d = new Date(onSendTab() ? sendY : viewY, (onSendTab() ? sendM : viewM) + step, 1);
    goMonth(d.getFullYear(), d.getMonth());
  }

  function goToday() {
    const now = parse(todayIso);
    weekExtra = new Set();
    if (currentTab === "week") setDate(todayIso);
    goMonth(now.getFullYear(), now.getMonth());
  }

  async function openMonths() {
    const y = onSendTab() ? sendY : viewY, m = onSendTab() ? sendM : viewM, now = parse(todayIso), waiting = pendingByMonth();
    let year = Math.min(Math.max(y, FIRST_YEAR), now.getFullYear());
    const choice = await ask("בחירת חודש", [], [{ label: "החודש הנוכחי", value: "now", primary: true }, { label: "ביטול", value: null }], (body, done) => {
      const years = document.createElement("div"), grid = document.createElement("div");
      years.id = "myears"; years.className = "opts"; grid.className = "mgrid";
      const draw = () => {
        const list = [];
        for (let v = now.getFullYear(); v >= FIRST_YEAR; v--) list.push([v, String(v)]);
        choices("myears", list, year, v => { year = v; draw(); });
        grid.textContent = "";
        HE_MONTHS.forEach((name, i) => {
          const b = document.createElement("button");
          b.className = "mcell" + (waiting[keyOf(year, i)] ? " waiting" : "");
          b.textContent = "(" + (i + 1) + ") " + name;
          b.setAttribute("aria-pressed", String(year === y && i === m));
          b.onclick = () => done({ y: year, m: i });
          grid.appendChild(b);
        });
      };
      body.append(years, grid);
      draw();
    });
    if (choice === "now") goToday();
    else if (choice) goMonth(choice.y, choice.m);
  }

  $("mname").onclick = openMonths;
  $("prev").onclick = () => moveMonth(-1);
  $("next").onclick = () => moveMonth(1);
  $("wprev").onclick = () => moveWeek(-1);
  $("wnext").onclick = () => moveWeek(1);
  $("minus").onclick = () => { hours = Math.max(0.5, hours - 0.5); renderHours(); };
  $("plus").onclick = () => { hours = Math.min(23.5, hours + 0.5); renderHours(); };
  document.querySelectorAll(".tabbar button").forEach(b => {
    b.onclick = () => {
      const tab = b.dataset.tab, at = Date.now(), twice = lastTap.tab === tab && at - lastTap.at < 400;
      lastTap = { tab, at };
      if (twice && tab !== "projects") goToday(); else showTab(tab);
    };
  });
  $("pshow").onclick = () => { showAll = !showAll; renderPicker(); };
  $("xnow").onclick = goToday;
  $("gear").onclick = () => showTab(currentTab === "settings" ? lastTab : "settings");
  $("sback").onclick = () => showTab(lastTab);
  $("change").onclick = () => { picking = true; render(); };
  $("dmarkbtn").onclick = () => markDay(selDate);
  $("psearch").oninput = renderPicker;
  $("asearch").oninput = renderProjectList;
  $("wsearch").oninput = renderWeekSearch;
  $("nsave").onclick = saveProject;
  ["nnum", "nname"].forEach(id => $(id).addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); saveProject(); }
  }));
  $("add").onclick = () => {
    if (!project(selNum)) return;
    entries.push({ id: newId(), date: selDate, projectNum: selNum, hours, created: new Date().toISOString() });
    persist();
    toast("נשמר: " + fmtH(hours) + " ש׳ · " + label(selNum));
    picking = true;
    render();
  };
  $("xall").onclick = () => {
    const { pending, selected } = exportState();
    unchecked = new Set(selected.length ? pending.map(r => r.id) : []);
    renderExport();
  };
  $("xsend").onclick = sendMail;
  $("xshare").onclick = shareFile;
  $("xdone").onclick = () => setLoaded(exportState().selected, new Date().toISOString());
  $("sconnect").onclick = connect;
  $("stest").onclick = testMail;
  $("bsave").onclick = saveBackupCopy;
  $("bload").onclick = () => $("bfile").click();
  $("bfile").onchange = e => { if (e.target.files[0]) importBackup(e.target.files[0]); e.target.value = ""; };
  ["surl", "skey"].forEach(id => { $(id).oninput = () => { $(id).dataset.dirty = "1"; }; });
  $("semp").onchange = e => { settings.employeeNum = e.target.value.trim(); persist(); render(); };
  $("smail").onchange = e => { settings.recipient = e.target.value.trim(); persist(); render(); };

  // Zoom is off on purpose: pinching also enlarges the tab bar. Text size is a setting instead.
  ["gesturestart", "gesturechange"].forEach(name => document.addEventListener(name, e => e.preventDefault(), { passive: false }));
  document.addEventListener("touchmove", e => { if (e.touches.length > 1) e.preventDefault(); }, { passive: false });
  darkQuery.addEventListener("change", applyLook);

  window.addEventListener("online", runBackup);
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { runBackup(); return; }
    const now = iso(new Date());
    if (now !== todayIso) {
      todayIso = now; setDate(now);
      sendY = viewY; sendM = viewM;
      render();
    }
    runBackup();
  });

  loadLocal();
  showTab(connected() ? "report" : "settings");
  runBackup();

  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
})();
