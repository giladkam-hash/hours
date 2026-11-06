(function () {
  const STORE_KEY = "hours.v1";
  const HE_MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];
  const HE_DAYS = ["א׳", "ב׳", "ג׳", "ד׳", "ה׳", "ו׳", "ש׳"];
  const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const SUGGESTIONS = 4;
  const DAY_TYPES = { eve: "ערב חג", holiday: "חג", vacation: "חופש", sick: "מחלה" };
  // Rides in the projects list, so aliases and day marks reach the backup without changing Code.gs.
  const META_NUM = "__app";

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
  let selDate = todayIso, selNum = null, hours = 1, picking = true, armed = null, aliasing = null, scrolledTo = null, editingId = null;
  let unchecked = new Set(), weekExtra = new Set(), sending = false, connecting = false, currentTab = "report", lastTab = "report";

  let projects = [], entries = [], aliases = {}, dayMarks = {};
  let settings = { scriptUrl: "", key: "", recipient: "", employeeNum: "", lastBackup: null, backupPending: false };
  let backupError = null, storageFailed = false, backupRunning = false, changes = 0, backupTimer = null;

  function toast(text) {
    const el = document.createElement("div");
    el.className = "toast"; el.setAttribute("role", "status"); el.textContent = text;
    document.body.appendChild(el);
    setTimeout(() => el.remove(), 2600);
  }

  /* Bottom sheet that resolves with the value of the button pressed. */
  function ask(title, lines, buttons) {
    return new Promise(resolve => {
      $("mtitle").textContent = title;
      const body = $("mbody"), box = $("mbtns");
      body.textContent = ""; box.textContent = "";
      lines.forEach(text => { const p = document.createElement("p"); p.textContent = text; body.appendChild(p); });
      buttons.forEach(b => {
        const el = document.createElement("button");
        el.className = b.primary ? "primary" : "secondary";
        el.textContent = b.label;
        el.onclick = () => { $("modal").hidden = true; resolve(b.value); };
        box.appendChild(el);
      });
      $("modal").hidden = false;
    });
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
      settings = Object.assign(settings, saved.settings);
    } catch (e) { storageFailed = true; }
  }

  function writeLocal() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({ projects, entries, aliases, dayMarks, settings }));
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

  const backupProjects = () => projects.concat({ num: META_NUM, name: JSON.stringify({ aliases, days: dayMarks }) });
  const upload = (url, key) => callScript(url, key, "save", { projects: backupProjects(), entries });

  /* Takes aliases and day marks from a backup. What is already on this device wins. */
  function mergeMeta(text) {
    let meta;
    try { meta = JSON.parse(text); } catch (e) { return; }
    Object.entries(meta.aliases || {}).forEach(([num, alias]) => {
      if (typeof alias === "string" && alias && !aliases[num]) aliases[num] = alias;
    });
    Object.entries(meta.days || {}).forEach(([date, type]) => {
      if (/^\d{4}-\d{2}-\d{2}$/.test(date) && Object.keys(DAY_TYPES).includes(type) && !dayMarks[date]) dayMarks[date] = type;
    });
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
  const monthKey = () => viewY + "-" + pad(viewM + 1);
  const monthEntries = () => entries.filter(e => e.date.startsWith(monthKey()));
  const norm = s => String(s || "").replace(/["'״׳]/g, "").toLowerCase();
  const matches = (p, q) => !q || p.num.includes(q) || norm(p.name).includes(q) || norm(aliases[p.num]).includes(q);
  const byNumDesc = (a, b) => (Number(b.num) || 0) - (Number(a.num) || 0);
  const hoursOn = date => entries.reduce((s, e) => (e.date === date ? s + e.hours : s), 0);

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
    if (date > todayIso) return "";
    if (h === 0) return date === todayIso ? "" : "missing";
    if (h < rule.min) return "low";
    if (h < rule.norm) return "short";
    return "ok";
  }

  function monthDays() {
    const count = new Date(viewY, viewM + 1, 0).getDate(), days = [];
    for (let d = 1; d <= count; d++) days.push(iso(new Date(viewY, viewM, d)));
    return days;
  }

  /* Workdays up to today that are missing a report or below the office hours, and days off that have hours. */
  function dayIssues() {
    const list = [];
    monthDays().filter(k => k <= todayIso).forEach(k => {
      const rule = dayRule(k), h = hoursOn(k);
      if (isRest(k) && h) list.push({ date: k, text: dayLabel(k) + ": מסומן " + dayMarkText(k) + " אבל יש " + fmtH(h) + " ש׳" });
      if (!rule) return;
      if (h === 0) list.push({ date: k, text: dayLabel(k) + ": אין דיווח" });
      else if (h < rule.min) list.push({ date: k, text: dayLabel(k) + ": " + fmtH(h) + " ש׳, מתחת למינימום " + rule.min });
      else if (h < rule.norm) list.push({ date: k, text: dayLabel(k) + ": " + fmtH(h) + " ש׳, מתחת ל-" + rule.norm + " של המשרד" });
    });
    return list;
  }

  /* The last day reported before this one must be fully included, then the most recent projects fill the rest. */
  function suggested(date) {
    const sorted = entries.slice().sort((a, b) => b.date.localeCompare(a.date) || String(b.created).localeCompare(String(a.created)));
    const earlier = sorted.filter(e => e.date < date);
    const lastDay = earlier.length ? earlier[0].date : null;
    const nums = [];
    earlier.filter(e => e.date === lastDay).forEach(e => { if (!nums.includes(e.projectNum)) nums.push(e.projectNum); });
    earlier.concat(sorted).forEach(e => {
      if (nums.length < SUGGESTIONS && !nums.includes(e.projectNum)) nums.push(e.projectNum);
    });
    return nums.map(project).filter(Boolean);
  }

  function monthRows() {
    const by = new Map();
    monthEntries().forEach(e => {
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

  async function markDay(date) {
    if (document.activeElement) document.activeElement.blur();
    const now = dayMarks[date] || "";
    const options = [["", "יום רגיל"], ["eve", "ערב חג (5 שעות)"], ["holiday", "חג"], ["vacation", "יום חופש"], ["sick", "יום מחלה"]];
    const choice = await ask("סימון " + dayLabel(date), ["בחג, בחופש ובמחלה לא נדרש דיווח."],
      options.map(([value, text]) => ({ label: text, value, primary: value === now })).concat({ label: "ביטול", value: null }));
    if (choice === null || choice === now) return;
    if (choice) dayMarks[date] = choice; else delete dayMarks[date];
    persist();
    render();
  }

  function saveAlias(num, text) {
    const alias = text.trim();
    const taken = alias && projects.find(p => p.num !== num && aliases[p.num] === alias);
    if (taken) { toast("הכינוי הזה כבר שייך לפרויקט " + taken.num); return; }
    if (alias) aliases[num] = alias; else delete aliases[num];
    aliasing = null;
    persist();
    render();
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
    b.children[1].textContent = p.num;
    b.onclick = onPick;
    return b;
  }

  // MARK: Header and days

  function renderHero() {
    $("mname").textContent = HE_MONTHS[viewM] + " " + viewY;
    const total = monthEntries().reduce((s, e) => s + e.hours, 0);
    const need = monthDays().filter(k => k <= todayIso).reduce((s, k) => { const r = dayRule(k); return s + (r ? r.norm : 0); }, 0);
    $("mtotal").textContent = fmtH(total);
    $("mneed").textContent = fmtH(need);
    $("mgaplabel").textContent = total >= need ? "עודף" : "חסר";
    $("mgap").textContent = fmtH(Math.abs(need - total));
  }

  function renderDays() {
    const box = $("days");
    box.textContent = "";
    monthDays().forEach(k => {
      const h = hoursOn(k), wd = parse(k).getDay();
      const b = document.createElement("button");
      b.className = ["day", dayState(k), k === todayIso ? "today" : "", k === selDate ? "sel" : ""].join(" ");
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
    const type = dayMarks[selDate];
    $("daybar").hidden = !isWorkday(selDate);
    $("dmark").textContent = !type ? "" : type === "eve" ? "ערב חג: נדרשות 5 שעות" : dayMarkText(selDate) + ": לא נדרש דיווח";
    $("dmarkbtn").textContent = type ? "שינוי" : "סימון חופש, מחלה או חג";
  }

  // MARK: Daily report

  function renderPicker() {
    const cur = project(selNum), show = !!cur && !picking;
    $("chosen").hidden = !show;
    $("picker").hidden = show;
    if (show) { $("cnum").textContent = cur.num; $("cname").textContent = label(cur.num); return; }
    const choose = p => () => { selNum = p.num; picking = false; $("psearch").value = ""; render(); };
    const raw = $("psearch").value.trim(), q = norm(raw), box = $("plist"), sug = $("sug");
    const recent = suggested(selDate);
    sug.textContent = "";
    sug.hidden = !!q || !recent.length;
    recent.forEach(p => sug.appendChild(projectButton(p, "sugbtn", choose(p))));
    box.textContent = "";
    if (q) {
      const hits = projects.filter(p => matches(p, q)).sort(byNumDesc);
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
    projects.filter(p => !recent.includes(p)).sort(byNumDesc).forEach(p => box.appendChild(projectButton(p, "pick", choose(p))));
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
        mark.className = "sentmark"; mark.textContent = "נשלח";
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
      el.className = "wday" + (dayMarks[k] ? " marked" : "");
      el.innerHTML = "<b></b><small></small>";
      el.children[0].textContent = HE_DAYS[parse(k).getDay()];
      el.children[1].textContent = dayMarkText(k) || shortDate(k);
      el.setAttribute("aria-label", "סימון " + dayLabel(k));
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
    projects.filter(p => matches(p, q)).sort(byNumDesc).slice(0, 8).forEach(p => {
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
      b.className = ["cell", dayState(k), k === todayIso ? "today" : ""].join(" ");
      b.innerHTML = '<span class="n"></span><span class="h"></span>';
      b.children[0].textContent = parse(k).getDate();
      b.children[1].textContent = h ? fmtH(h) : dayMarkText(k);
      b.children[1].classList.toggle("t", !h);
      b.setAttribute("aria-label", dayLabel(k) + (dayMarks[k] ? ", " + dayMarkText(k) : "") + (h ? ", " + fmtH(h) + " שעות" : ""));
      b.onclick = () => { selDate = k; scrolledTo = null; showTab("report"); };
      box.appendChild(b);
    });
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
    box.textContent = "";
    if (!pending.length && !done.length) box.appendChild(paragraph("אין דיווחים בחודש הזה.", "empty"));
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
    monthEntries().forEach(e => { byDay[e.date] = (byDay[e.date] || 0) + e.hours; });
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

  function aliasEditor(p) {
    const box = document.createElement("div");
    box.className = "aliasedit";
    box.innerHTML = '<input type="text" maxlength="40" autocomplete="off" placeholder="כינוי קצר" aria-label="כינוי לפרויקט"><button class="linkbtn">שמירה</button>';
    const input = box.children[0];
    input.value = aliases[p.num] || "";
    input.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); saveAlias(p.num, input.value); } });
    box.children[1].onclick = () => saveAlias(p.num, input.value);
    return box;
  }

  function renderProjectList() {
    const q = norm($("asearch").value.trim()), box = $("pall");
    box.textContent = "";
    $("pcount").textContent = projects.length + " פרויקטים";
    const used = new Set(entries.map(e => e.projectNum));
    const list = projects.filter(p => matches(p, q)).sort(byNumDesc);
    if (!list.length) { box.appendChild(paragraph("לא נמצאו פרויקטים.", "empty")); return; }
    list.forEach(p => {
      const row = document.createElement("div");
      row.className = "prow";
      row.innerHTML = '<span class="pname"><span></span><small class="porig"></small></span><span class="pnum"></span><button class="linkbtn"></button>';
      row.querySelector(".pname span").textContent = label(p.num);
      row.querySelector(".porig").textContent = aliases[p.num] ? p.name : "";
      row.querySelector(".pnum").textContent = p.num;
      const edit = row.querySelector(".linkbtn");
      edit.textContent = aliases[p.num] ? "שינוי כינוי" : "כינוי";
      edit.setAttribute("aria-label", "כינוי ל" + p.name);
      edit.onclick = () => { aliasing = aliasing === p.num ? null : p.num; renderProjectList(); };
      if (!used.has(p.num)) {
        const del = document.createElement("button");
        del.className = "del";
        del.textContent = armed === p.num ? "למחוק?" : "×";
        del.setAttribute("aria-label", "מחיקת " + p.name);
        del.onclick = () => {
          if (armed !== p.num) { armed = p.num; renderProjectList(); return; }
          armed = null;
          projects = projects.filter(x => x.num !== p.num);
          delete aliases[p.num];
          if (selNum === p.num) { selNum = null; picking = true; }
          persist();
          render();
        };
        row.appendChild(del);
      }
      box.appendChild(row);
      if (aliasing === p.num) {
        row.appendChild(aliasEditor(p));
        row.querySelector(".aliasedit input").focus();
      }
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
    renderHero(); renderDays(); renderDayBar(); renderPicker(); renderHours(); renderLog();
    if (document.activeElement === null || !document.activeElement.closest("#wrows")) renderWeek();
    renderCalendar(); renderSum(); renderExport(); renderProjectList(); renderSettings(); renderStatus();
  }

  function showTab(name) {
    if (name !== "settings") lastTab = name;
    currentTab = name;
    ["report", "week", "month", "projects", "settings"].forEach(t => { $("tab-" + t).hidden = t !== name; });
    document.querySelectorAll(".tabbar button").forEach(b => {
      if (b.dataset.tab === name) b.setAttribute("aria-current", "page"); else b.removeAttribute("aria-current");
    });
    $("monthbar").hidden = name !== "report" && name !== "month";
    armed = null; aliasing = null; editingId = null;
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
    selDate = issues[0].date; scrolledTo = null;
    showTab("report");
    return false;
  }

  function exportFile(selected) {
    return { name: "master-plan-hours-" + monthKey() + ".xlsx", bytes: HoursXlsx.file(settings.employeeNum, selected) };
  }

  function base64(bytes) {
    let text = "";
    bytes.forEach(b => { text += String.fromCharCode(b); });
    return btoa(text);
  }

  async function sendMail() {
    if (!await reviewDays()) return;
    const { selected } = exportState();
    const title = HE_MONTHS[viewM] + " " + viewY, file = exportFile(selected);
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

  function moveMonth(step) {
    let y = viewY, m = viewM + step;
    if (m < 0) { m = 11; y -= 1; }
    if (m > 11) { m = 0; y += 1; }
    setDate(iso(new Date(y, m, 1)));
    render();
  }

  $("prev").onclick = () => moveMonth(-1);
  $("next").onclick = () => moveMonth(1);
  $("wprev").onclick = () => moveWeek(-1);
  $("wnext").onclick = () => moveWeek(1);
  $("minus").onclick = () => { hours = Math.max(0.5, hours - 0.5); renderHours(); };
  $("plus").onclick = () => { hours = Math.min(23.5, hours + 0.5); renderHours(); };
  document.querySelectorAll(".tabbar button").forEach(b => { b.onclick = () => showTab(b.dataset.tab); });
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
  $("semp").onchange = e => { settings.employeeNum = e.target.value.trim(); writeLocal(); render(); };
  $("smail").onchange = e => { settings.recipient = e.target.value.trim(); writeLocal(); render(); };

  window.addEventListener("online", runBackup);
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { runBackup(); return; }
    const now = iso(new Date());
    if (now !== todayIso) { todayIso = now; setDate(now); render(); }
    runBackup();
  });

  loadLocal();
  showTab(connected() ? "report" : "settings");
  runBackup();

  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
})();
