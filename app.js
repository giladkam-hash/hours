(function () {
  const STORE_KEY = "hours.v1";
  const HE_MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];
  const HE_DAYS = ["א׳", "ב׳", "ג׳", "ד׳", "ה׳", "ו׳", "ש׳"];
  const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  const $ = id => document.getElementById(id);
  const pad = n => String(n).padStart(2, "0");
  const iso = d => d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  const fmtH = h => String(Math.round(h * 100) / 100);
  const shortDate = k => { const [, m, d] = k.split("-"); return d + "/" + m; };
  const newId = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now() + "-" + Math.random().toString(16).slice(2));

  const today = new Date(), todayIso = iso(today);
  let viewY = today.getFullYear(), viewM = today.getMonth();
  let selDate = todayIso, selNum = null, hours = 1, picking = true, armed = null, scrolledTo = null;
  let unchecked = new Set(), sending = false, connecting = false, currentTab = "report";

  let projects = [], entries = [];
  let settings = { scriptUrl: "", key: "", recipient: "", employeeNum: "", lastBackup: null, backupPending: false };
  let backupError = null, storageFailed = false, backupRunning = false, changes = 0, backupTimer = null;

  function toast(text) {
    const el = document.createElement("div");
    el.className = "toast"; el.setAttribute("role", "status"); el.textContent = text;
    document.body.appendChild(el);
    setTimeout(() => el.remove(), 2600);
  }

  // MARK: Storage

  function loadLocal() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (!raw) return;
      const saved = JSON.parse(raw);
      projects = saved.projects || [];
      entries = saved.entries || [];
      settings = Object.assign(settings, saved.settings);
    } catch (e) { storageFailed = true; }
  }

  function writeLocal() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({ projects, entries, settings }));
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

  const upload = (url, key) => callScript(url, key, "save", { projects, entries });

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

  /* Merges so that neither side loses anything. Returns how many reports came from the backup. */
  function merge(remoteProjects, remoteEntries) {
    remoteProjects.forEach(p => { if (!project(p.num)) projects.push(p); });
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
  const monthKey = () => viewY + "-" + pad(viewM + 1);
  const monthEntries = () => entries.filter(e => e.date.startsWith(monthKey()));
  const norm = s => String(s || "").replace(/["'״׳]/g, "").toLowerCase();
  const matches = (p, q) => !q || p.num.includes(q) || norm(p.name).includes(q);
  const byNumDesc = (a, b) => (Number(b.num) || 0) - (Number(a.num) || 0);

  function recentProjects() {
    const seen = [];
    entries.slice().sort((a, b) => String(b.created).localeCompare(String(a.created))).forEach(e => {
      if (!seen.includes(e.projectNum)) seen.push(e.projectNum);
    });
    return seen.map(project).filter(Boolean).slice(0, 5);
  }

  function monthRows() {
    const by = new Map();
    monthEntries().forEach(e => {
      const loaded = !!e.loadedAt, id = e.date + "|" + e.projectNum + "|" + (loaded ? 1 : 0);
      if (!by.has(id)) by.set(id, { id, date: e.date, num: e.projectNum, name: nameOf(e.projectNum), hours: 0, entryIds: [], loaded });
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

  // MARK: Rendering

  function renderDays() {
    $("mname").textContent = HE_MONTHS[viewM] + " " + viewY;
    const count = new Date(viewY, viewM + 1, 0).getDate(), box = $("days");
    const byDay = {};
    monthEntries().forEach(e => { byDay[e.date] = (byDay[e.date] || 0) + e.hours; });
    box.textContent = "";
    for (let d = 1; d <= count; d++) {
      const date = new Date(viewY, viewM, d), k = iso(date), wd = date.getDay();
      const b = document.createElement("button");
      b.className = "day" + (wd >= 5 ? " weekend" : "") + (k === todayIso ? " today" : "") + (k === selDate ? " sel" : "");
      if (k === selDate) b.setAttribute("aria-current", "date");
      b.innerHTML = '<small></small><span class="n"></span><span class="h"></span>';
      b.children[0].textContent = HE_DAYS[wd];
      b.children[1].textContent = d;
      b.children[2].textContent = byDay[k] ? fmtH(byDay[k]) : "";
      b.onclick = () => { selDate = k; render(); };
      box.appendChild(b);
      if (k === selDate && scrolledTo !== k) {
        scrolledTo = k;
        requestAnimationFrame(() => b.scrollIntoView({ inline: "center", block: "nearest" }));
      }
    }
    $("mtotal").textContent = fmtH(monthEntries().reduce((s, e) => s + e.hours, 0));
  }

  function pickRow(p) {
    const b = document.createElement("button");
    b.className = "pick";
    b.innerHTML = '<span class="pnum"></span><span class="pname"></span>';
    b.children[0].textContent = p.num;
    b.children[1].textContent = p.name;
    b.onclick = () => { selNum = p.num; picking = false; $("psearch").value = ""; render(); };
    return b;
  }

  function paragraph(text, className) {
    const el = document.createElement("p");
    el.className = className; el.textContent = text;
    return el;
  }

  function listLabel(text) {
    const el = document.createElement("span");
    el.className = "label"; el.textContent = text;
    return el;
  }

  function renderPicker() {
    const cur = project(selNum), show = !!cur && !picking;
    $("chosen").hidden = !show;
    $("picker").hidden = show;
    if (show) { $("cnum").textContent = cur.num; $("cname").textContent = cur.name; return; }
    const raw = $("psearch").value.trim(), q = norm(raw), box = $("plist");
    box.textContent = "";
    if (q) {
      const hits = projects.filter(p => matches(p, q)).sort(byNumDesc);
      hits.forEach(p => box.appendChild(pickRow(p)));
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
    const recent = recentProjects();
    if (recent.length) {
      box.appendChild(listLabel("לאחרונה"));
      recent.forEach(p => box.appendChild(pickRow(p)));
      box.appendChild(listLabel("כל הפרויקטים"));
    }
    projects.filter(p => !recent.includes(p)).sort(byNumDesc).forEach(p => box.appendChild(pickRow(p)));
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
      row.innerHTML = '<span class="pnum"></span><span class="pname"></span>';
      row.children[0].textContent = p.num;
      row.children[1].textContent = p.name;
      if (!used.has(p.num)) {
        const del = document.createElement("button");
        del.className = "del";
        del.textContent = armed === p.num ? "למחוק?" : "×";
        del.setAttribute("aria-label", "מחיקת " + p.name);
        del.onclick = () => {
          if (armed !== p.num) { armed = p.num; renderProjectList(); return; }
          armed = null;
          projects = projects.filter(x => x.num !== p.num);
          if (selNum === p.num) { selNum = null; picking = true; }
          persist();
          render();
        };
        row.appendChild(del);
      }
      box.appendChild(row);
    });
  }

  function renderHours() {
    $("hours").textContent = fmtH(hours);
    const quick = $("quick");
    quick.textContent = "";
    [1, 2, 3, 4, 6, 8, 9].forEach(v => {
      const b = document.createElement("button");
      b.className = "chip" + (v === hours ? " on" : "");
      b.textContent = v;
      b.onclick = () => { hours = v; renderHours(); };
      quick.appendChild(b);
    });
    const d = new Date(selDate + "T00:00");
    $("add").textContent = "הוספה ליום " + HE_DAYS[d.getDay()] + " " + pad(d.getDate()) + "/" + pad(d.getMonth() + 1);
    $("add").disabled = !project(selNum);
  }

  function renderLog() {
    const d = new Date(selDate + "T00:00");
    $("dayh").textContent = "דיווחים ליום " + HE_DAYS[d.getDay()] + " " + pad(d.getDate()) + "/" + pad(d.getMonth() + 1);
    const list = entries.filter(e => e.date === selDate), box = $("log");
    box.textContent = "";
    $("dtotal").textContent = list.length ? fmtH(list.reduce((s, e) => s + e.hours, 0)) + " ש׳" : "";
    if (!list.length) { box.appendChild(paragraph("אין דיווחים ליום הזה.", "empty")); return; }
    list.forEach(e => {
      const row = document.createElement("div");
      row.className = "item";
      row.innerHTML = '<span class="p"></span><span class="hh"></span><button class="del" aria-label="מחיקת הדיווח">×</button>';
      row.children[0].textContent = nameOf(e.projectNum);
      row.children[1].textContent = fmtH(e.hours);
      row.children[2].onclick = () => { entries = entries.filter(x => x.id !== e.id); persist(); render(); };
      box.appendChild(row);
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
      tr.children[0].textContent = nameOf(num);
      tr.children[1].textContent = fmtH(h);
      table.appendChild(tr);
    });
    box.appendChild(table);
  }

  function exportProblems(selected) {
    const list = [];
    if (!settings.employeeNum) list.push("חסר מספר עובד. ממלאים אותו בלשונית הגדרות.");
    selected.filter(r => r.hours >= 24).forEach(r => {
      list.push("ב-" + shortDate(r.date) + " רשומות " + fmtH(r.hours) + " שעות על " + r.name + ". הקובץ תומך בפחות מ-24 שעות בשורה.");
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
      row.children[2].textContent = r.name;
      row.children[3].textContent = fmtH(r.hours);
      box.appendChild(row);
    });
    done.forEach(r => {
      const row = document.createElement("div");
      row.className = "xrow done";
      row.innerHTML = '<span class="ok">נקלט</span><span class="d"></span><span class="p"></span><span class="hh"></span><button class="linkbtn">החזרה</button>';
      row.children[1].textContent = shortDate(r.date);
      row.children[2].textContent = r.name;
      row.children[3].textContent = fmtH(r.hours);
      row.children[4].setAttribute("aria-label", "החזרת " + r.name + " " + shortDate(r.date) + " לרשימה");
      row.children[4].onclick = () => setLoaded([r], null);
      box.appendChild(row);
    });

    const problems = exportProblems(selected), messages = problems.slice();
    if (!settings.recipient) messages.push("חסרה כתובת מייל. ממלאים אותה בלשונית הגדרות.");
    const byDay = {};
    monthEntries().forEach(e => { byDay[e.date] = (byDay[e.date] || 0) + e.hours; });
    Object.keys(byDay).sort().forEach(k => {
      if (byDay[k] > 12) messages.push("ב-" + shortDate(k) + " רשומות " + fmtH(byDay[k]) + " שעות. כדאי לוודא שאין דיווח כפול.");
    });
    const warn = $("xwarn");
    warn.textContent = "";
    messages.forEach(text => warn.appendChild(paragraph(text, "warn")));

    const canExport = selected.length > 0 && !problems.length && !sending;
    $("xsend").textContent = sending ? "שולח…" : "שליחה במייל (" + selected.length + " שורות)";
    $("xsend").disabled = !canExport || !settings.recipient || !connected();
    $("xshare").disabled = !canExport;
    $("xdone").textContent = "סימון " + selected.length + " שורות כנקלטו";
    $("xdone").disabled = !selected.length;
    $("xall").textContent = selected.length ? "ניקוי הבחירה" : "בחירת הכל";
    $("xall").hidden = !pending.length;
    $("xnote").textContent = (settings.recipient ? "המייל נשלח אל " + settings.recipient + ". " : "")
      + "במאסטר פלן: שעות לפרויקטים, תצוגת רשימה, קליטת שעות מקובץ אקסל. אחרי הקליטה מסמנים כאן את השורות, והן לא ייכנסו לקובץ הבא.";
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
    else if (!connected()) text = "הגיבוי לא מחובר. בלי חיבור הנתונים נשמרים רק במכשיר הזה, ורשימת הפרויקטים ריקה. מחברים אותו בלשונית הגדרות.";
    else if (backupError) text = backupError;
    banner.textContent = text;
    banner.hidden = !text;
    if (connected() && !connecting && !$("sstatus").classList.contains("warn")) {
      const when = new Date(settings.lastBackup);
      setConnectStatus((settings.backupPending ? "יש שינויים שממתינים לגיבוי. " : "") + "גיבוי אחרון: " + shortDate(iso(when)) + " " + pad(when.getHours()) + ":" + pad(when.getMinutes()), false);
    }
  }

  function render() {
    renderDays(); renderPicker(); renderProjectList(); renderHours(); renderLog();
    renderSum(); renderExport(); renderSettings(); renderStatus();
  }

  function showTab(name) {
    currentTab = name;
    ["report", "projects", "export", "settings"].forEach(t => { $("tab-" + t).hidden = t !== name; });
    document.querySelectorAll(".tabbar button").forEach(b => {
      if (b.dataset.tab === name) b.setAttribute("aria-current", "page"); else b.removeAttribute("aria-current");
    });
    $("monthbar").hidden = name === "projects" || name === "settings";
    armed = null;
    render();
    window.scrollTo(0, 0);
  }

  // MARK: Export actions

  function exportFile(selected) {
    return { name: "master-plan-hours-" + monthKey() + ".xlsx", bytes: HoursXlsx.file(settings.employeeNum, selected) };
  }

  function base64(bytes) {
    let text = "";
    bytes.forEach(b => { text += String.fromCharCode(b); });
    return btoa(text);
  }

  async function sendMail() {
    const { selected } = exportState();
    const title = HE_MONTHS[viewM] + " " + viewY, file = exportFile(selected);
    sending = true; renderExport();
    try {
      await callScript(settings.scriptUrl, settings.key, "mail", {
        to: settings.recipient, subject: "דיווח שעות למאסטר פלן - " + title,
        text: "מצורף קובץ שעות ל" + title + ", " + selected.length + " שורות.",
        filename: file.name, file: base64(file.bytes),
      });
      toast("המייל נשלח אל " + settings.recipient);
    } catch (e) {
      toast(e.message);
    }
    sending = false; render();
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

  async function shareFile() {
    const { selected } = exportState(), file = exportFile(selected);
    const blob = new File([file.bytes], file.name, { type: XLSX_TYPE });
    if (navigator.canShare && navigator.canShare({ files: [blob] })) {
      try { await navigator.share({ files: [blob] }); } catch (e) { if (e.name !== "AbortError") toast("השיתוף נכשל"); }
      return;
    }
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = file.name;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 10000);
  }

  // MARK: Events

  function moveMonth(step) {
    viewM += step;
    if (viewM < 0) { viewM = 11; viewY -= 1; }
    if (viewM > 11) { viewM = 0; viewY += 1; }
    selDate = iso(new Date(viewY, viewM, 1));
    render();
  }

  $("prev").onclick = () => moveMonth(-1);
  $("next").onclick = () => moveMonth(1);
  $("minus").onclick = () => { hours = Math.max(0.5, hours - 0.5); renderHours(); };
  $("plus").onclick = () => { hours = Math.min(23.5, hours + 0.5); renderHours(); };
  document.querySelectorAll(".tabbar button").forEach(b => { b.onclick = () => showTab(b.dataset.tab); });
  $("change").onclick = () => { picking = true; render(); $("psearch").focus(); };
  $("psearch").oninput = renderPicker;
  $("asearch").oninput = renderProjectList;
  $("nsave").onclick = saveProject;
  ["nnum", "nname"].forEach(id => $(id).addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); saveProject(); }
  }));
  $("add").onclick = () => {
    if (!project(selNum)) return;
    entries.push({ id: newId(), date: selDate, projectNum: selNum, hours, created: new Date().toISOString() });
    persist();
    toast("נשמר: " + fmtH(hours) + " ש׳ · " + nameOf(selNum));
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
  ["surl", "skey"].forEach(id => { $(id).oninput = () => { $(id).dataset.dirty = "1"; }; });
  $("semp").onchange = e => { settings.employeeNum = e.target.value.trim(); writeLocal(); render(); };
  $("smail").onchange = e => { settings.recipient = e.target.value.trim(); writeLocal(); render(); };

  window.addEventListener("online", runBackup);
  document.addEventListener("visibilitychange", runBackup);

  loadLocal();
  showTab(connected() ? "report" : "settings");
  runBackup();

  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
})();
