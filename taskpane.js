/* global Office, Word */

let rows = [
  { id: 1, field: "", value: "" },
  { id: 2, field: "", value: "" },
  { id: 3, field: "", value: "" },
  { id: 4, field: "", value: "" },
];

let selectedRowId = null;
let saveTimer = null;
let statusTimer = null;

function $(id) {
  return document.getElementById(id);
}

function normalizeKey(s) {
  return (s || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "_")
    .replace(/[^A-Z0-9_]/g, "");
}

function showStatus(msg, autohide) {
  const el = $("status");
  if (!el) return;

  el.textContent = msg || "";
  el.style.display = msg ? "block" : "none";

  if (statusTimer) {
    clearTimeout(statusTimer);
    statusTimer = null;
  }

  if (autohide && msg) {
    statusTimer = window.setTimeout(() => {
      el.style.display = "none";
      el.textContent = "";
    }, 4000);
  }
}

/* =========================
   TABLE
   ========================= */

function getMap() {
  const m = new Map();
  for (const r of rows) {
    const k = normalizeKey(r.field);
    if (k) m.set(k, (r.value == null ? "" : r.value));
  }
  return m;
}

function addRow() {
  const nextId = Math.max(...rows.map((r) => r.id), 0) + 1;
  rows.push({ id: nextId, field: "", value: "" });
  if (selectedRowId == null) selectedRowId = nextId;
  renderRows();
  triggerAutoSave();
}

function deleteRow(id) {
  if (rows.length <= 1) {
    showStatus("⚠ Mora ostati bar jedan red u tabeli", true);
    return;
  }
  rows = rows.filter((r) => r.id !== id);
  if (selectedRowId === id) {
    selectedRowId = rows[0] ? rows[0].id : null;
  }
  renderRows();
  triggerAutoSave();
}

function updateRow(id, key, val) {
  rows = rows.map((r) => (r.id === id ? { ...r, [key]: val } : r));
  triggerAutoSave();
}

function getSelectedKey() {
  if (selectedRowId == null) return null;
  const r = rows.find((x) => x.id === selectedRowId);
  const key = normalizeKey((r && r.field) || "");
  return key || null;
}

function renderRows() {
  const container = $("rows");
  if (!container) return;
  container.innerHTML = "";

  rows.forEach((r) => {
    const row = document.createElement("div");
    row.className = "row";

    if (selectedRowId === r.id) {
      row.style.background = "#eff6ff";
    }

    const c1 = document.createElement("div");
    c1.className = "cell";
    const i1 = document.createElement("input");
    i1.value = r.field;
    i1.placeholder = "Unesite polje...";
    i1.addEventListener("focus", () => {
      if (selectedRowId !== r.id) {
        selectedRowId = r.id;
        renderRows();
      }
    });
    i1.addEventListener("input", (e) => updateRow(r.id, "field", e.target.value));
    c1.appendChild(i1);

    const c2 = document.createElement("div");
    c2.className = "cell";
    const i2 = document.createElement("input");
    i2.value = r.value;
    i2.placeholder = "Unesite odgovor...";
    i2.addEventListener("focus", () => {
      if (selectedRowId !== r.id) {
        selectedRowId = r.id;
        renderRows();
      }
    });
    i2.addEventListener("input", (e) => updateRow(r.id, "value", e.target.value));
    c2.appendChild(i2);

    const c3 = document.createElement("div");
    c3.className = "del";
    const b = document.createElement("button");
    b.textContent = "×";
    b.title = "Obriši red (Delete)";
    b.addEventListener("click", (ev) => {
      ev.stopPropagation();
      deleteRow(r.id);
    });
    c3.appendChild(b);

    row.appendChild(c1);
    row.appendChild(c2);
    row.appendChild(c3);
    container.appendChild(row);
  });
}

/* =========================
   TAG encoding: KEY|TYPE|FMT
   TYPE: T (text), D (date), N (number)
   ========================= */

function makeTag(key, type, fmt) {
  return `${key}|${type}|${fmt}`;
}

function parseTag(tag) {
  const parts = (tag || "").split("|");
  if (parts.length < 3) return null;

  const key = normalizeKey(parts[0]);
  const type = parts[1];
  const fmt = parts.slice(2).join("|");

  if (!key) return null;
  if (type !== "T" && type !== "D" && type !== "N") return null;

  return { key, type, fmt };
}

/* =========================
   FORMAT OPTIONS + MODAL
   ========================= */

const FORMAT_OPTIONS = {
  T: [
    { value: "AS_IS", label: "Kao uneto" },
    { value: "UPPER", label: "VELIKA SLOVA" },
    { value: "LOWER", label: "mala slova" },
    { value: "SENTENCE", label: "Rečenica (prvo veliko)" },
  ],
  D: [
    { value: "DD_MM_YYYY", label: "dd.mm.yyyy" },
    { value: "DD", label: "dan (dd)" },
    { value: "MM", label: "mesec broj (mm)" },
    { value: "YYYY", label: "godina (yyyy)" },
    { value: "MONTH_TEXT", label: "mesec tekst (januar...)" },
    { value: "MONTH_TEXT_YEAR", label: "April 2025" },
    { value: "DD_MONTH_TEXT_YYYY", label: "15. april 2025" },
  ],
  N: [
    { value: "PLAIN_0", label: "broj (0 dec)" },
    { value: "PLAIN_2", label: "broj (2 dec)" },
    { value: "CUR_EUR_2", label: "EUR (2 dec)" },
    { value: "CUR_RSD_0", label: "RSD (0 dec)" },
    { value: "CUR_USD_2", label: "USD (2 dec)" },
  ],
};

function showOverlay(show) {
  const el = $("modalOverlay");
  if (!el) return;
  el.style.display = show ? "flex" : "none";
}

function fillFormatOptions(type) {
  const sel = $("modalFormat");
  if (!sel) return;
  sel.innerHTML = "";

  const opts = FORMAT_OPTIONS[type] || [];
  for (const o of opts) {
    const opt = document.createElement("option");
    opt.value = o.value;
    opt.textContent = o.label;
    sel.appendChild(opt);
  }
}

async function openInsertModal(key) {
  const nameEl = $("modalFieldName");
  const typeEl = $("modalType");
  const fmtEl = $("modalFormat");
  const okBtn = $("modalOk");
  const cancelBtn = $("modalCancel");

  if (!nameEl || !typeEl || !fmtEl || !okBtn || !cancelBtn) {
    showStatus("⚠ Modal elementi nisu pronađeni u HTML-u", true);
    return null;
  }

  nameEl.textContent = key;

  typeEl.value = "T";
  fillFormatOptions("T");

  const onTypeChange = () => fillFormatOptions(typeEl.value);
  typeEl.addEventListener("change", onTypeChange);

  showOverlay(true);

  return await new Promise((resolve) => {
    const cleanup = () => {
      showOverlay(false);
      typeEl.removeEventListener("change", onTypeChange);
      okBtn.removeEventListener("click", onOk);
      cancelBtn.removeEventListener("click", onCancel);
    };

    const onOk = () => {
      const type = typeEl.value;
      const fmt = fmtEl.value;
      cleanup();
      resolve({ type, fmt });
    };

    const onCancel = () => {
      cleanup();
      resolve(null);
    };

    okBtn.addEventListener("click", onOk);
    cancelBtn.addEventListener("click", onCancel);
  });
}

/* ===== Confirm modal (Today date) ===== */

function showConfirmOverlay(show) {
  const el = $("confirmOverlay");
  if (!el) return;
  el.style.display = show ? "flex" : "none";
}

async function confirmToday(key) {
  const textEl = $("confirmText");
  const yesBtn = $("confirmYes");
  const noBtn = $("confirmNo");

  if (!textEl || !yesBtn || !noBtn) {
    // fallback: confirm()
    return window.confirm(`Polje "${key}" je prazno u tabeli.\nDa li je to današnji datum?`);
  }

  textEl.textContent = `Polje "${key}" je prazno u tabeli.\nDa li je to današnji datum?`;
  showConfirmOverlay(true);

  return await new Promise((resolve) => {
    const cleanup = () => {
      showConfirmOverlay(false);
      yesBtn.removeEventListener("click", onYes);
      noBtn.removeEventListener("click", onNo);
    };

    const onYes = () => { cleanup(); resolve(true); };
    const onNo = () => { cleanup(); resolve(false); };

    yesBtn.addEventListener("click", onYes);
    noBtn.addEventListener("click", onNo);
  });
}

/* =========================
   FORMATTING
   ========================= */

function toSentenceCase(s) {
  const t = (s || "").trim();
  if (!t) return "";
  return t.charAt(0).toUpperCase() + t.slice(1).toLowerCase();
}

function formatText(raw, fmt) {
  const v = raw == null ? "" : String(raw);
  switch (fmt) {
    case "UPPER": return v.toUpperCase();
    case "LOWER": return v.toLowerCase();
    case "SENTENCE": return toSentenceCase(v);
    case "AS_IS":
    default: return v;
  }
}

function parseDateLoose(raw) {
  const s = (raw || "").trim();
  if (!s) return null;

  const m1 = s.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})$/);
  if (m1) {
    const dd = Number(m1[1]);
    const mm = Number(m1[2]);
    const yyyy = Number(m1[3]);
    const d = new Date(yyyy, mm - 1, dd);
    if (!isNaN(d.getTime())) return d;
  }

  const m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m2) {
    const yyyy = Number(m2[1]);
    const mm = Number(m2[2]);
    const dd = Number(m2[3]);
    const d = new Date(yyyy, mm - 1, dd);
    if (!isNaN(d.getTime())) return d;
  }

  const d3 = new Date(s);
  if (!isNaN(d3.getTime())) return d3;

  return null;
}

const MONTHS_SR_LAT = [
  "januar","februar","mart","april","maj","jun","jul","avgust","septembar","oktobar","novembar","decembar"
];

function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatDate(raw, fmt) {
  const d = parseDateLoose(raw);
  if (!d) return raw == null ? "" : String(raw);

  const dd = d.getDate();
  const mm = d.getMonth() + 1;
  const yyyy = d.getFullYear();
  const m = MONTHS_SR_LAT[mm - 1] || "";

  switch (fmt) {
    case "DD": return pad2(dd);
    case "MM": return pad2(mm);
    case "YYYY": return String(yyyy);
    case "MONTH_TEXT": return m;
    case "MONTH_TEXT_YEAR": {
      const cap = m ? m.charAt(0).toUpperCase() + m.slice(1) : "";
      return `${cap} ${yyyy}`;
    }
    case "DD_MONTH_TEXT_YYYY":
      return `${dd}. ${m} ${yyyy}`;
    case "DD_MM_YYYY":
    default:
      return `${pad2(dd)}.${pad2(mm)}.${yyyy}`;
  }
}

function formatNumber(raw, fmt) {
  const s = (raw == null ? "" : String(raw)).toString().trim();
  if (!s) return "";

  const norm = s.replace(/\s/g, "").replace(",", ".");
  const n = Number(norm);
  if (isNaN(n)) return raw == null ? "" : String(raw);

  const m = fmt.match(/^(PLAIN|CUR)_(EUR|RSD|USD)?_(\d)$/);
  if (!m) return raw == null ? "" : String(raw);

  const kind = m[1];
  const cur = m[2] || "";
  const dec = Number(m[3]);

  const formatted = n.toLocaleString("sr-RS", {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec,
  });

  if (kind === "CUR") {
    const sym = cur === "EUR" ? "€" : (cur === "USD" ? "$" : (cur === "RSD" ? "RSD" : ""));
    return sym ? `${formatted} ${sym}` : formatted;
  }

  return formatted;
}

function applyFormat(raw, type, fmt) {
  if (type === "T") return formatText(raw, fmt);
  if (type === "D") return formatDate(raw, fmt);
  return formatNumber(raw, fmt);
}

function todayRawDDMMYYYY() {
  const d = new Date();
  const dd = pad2(d.getDate());
  const mm = pad2(d.getMonth() + 1);
  const yyyy = d.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

/* =========================
   PERSISTENCE (Custom XML)
   ========================= */

const XML_NS = "biroa/fields/v2";

function escapeXml(s) {
  return (s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

async function saveToDocument() {
  try {
    await Word.run(async (context) => {
      const parts = context.document.customXmlParts.getByNamespace(XML_NS);
      parts.load("items");
      await context.sync();
      parts.items.forEach((p) => p.delete());
      await context.sync();

      let xml = `<BiroA xmlns="${XML_NS}">`;
      for (const r of rows) {
        const key = normalizeKey(r.field);
        if (!key) continue;
        xml += `<field name="${escapeXml(key)}">`;
        xml += `<value>${escapeXml(r.value || "")}</value>`;
        xml += `</field>`;
      }
      xml += `</BiroA>`;

      context.document.customXmlParts.add(xml);
      await context.sync();
    });
  } catch (e) {
    console.error("Save failed:", e);
    // auto-save: ne smaraj korisnika
  }
}

async function loadFromDocument() {
  try {
    await Word.run(async (context) => {
      const parts = context.document.customXmlParts.getByNamespace(XML_NS);
      parts.load("items");
      await context.sync();

      if (!parts.items.length) return;

      const xmlRes = parts.items[0].getXml();
      await context.sync();

      const xml = (xmlRes && xmlRes.value) ? xmlRes.value : "";
      if (!xml) return;

      const doc = new DOMParser().parseFromString(xml, "text/xml");
      const fields = Array.from(doc.getElementsByTagName("field"));

      const newRows = [];
      let idx = 1;

      for (const f of fields) {
        const name = f.getAttribute("name") || "";
        const valueEl = f.getElementsByTagName("value")[0];
        const value = valueEl ? (valueEl.textContent || "") : "";
        if (!name) continue;
        newRows.push({ id: idx++, field: name, value });
      }

      if (newRows.length) {
        rows = newRows;
        selectedRowId = rows[0] ? rows[0].id : null;
        renderRows();
        showStatus(`✓ Učitano ${rows.length} polja iz dokumenta`, true);
      }
    });
  } catch (e) {
    console.error("Load failed:", e);
    showStatus(`⚠ Greška pri učitavanju: ${e && e.message ? e.message : String(e)}`, true);
  }
}

function triggerAutoSave() {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => { saveToDocument(); }, 1500);
}

/* =========================
   CSV IMPORT/EXPORT
   ========================= */

function exportCSV() {
  const lines = rows.map((r) => `${r.field},${r.value}`);
  const csv = lines.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "ba-polja.csv";
  a.click();

  URL.revokeObjectURL(url);
  showStatus("✓ CSV eksportovan", true);
}

function importCSV() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".csv,text/csv";

  input.onchange = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    const text = await file.text();
    const lines = text.trim().split("\n");
    const newRows = [];
    let id = 1;

    for (const line of lines) {
      const parts = line.split(",");
      const field = (parts[0] || "").trim();
      const value = (parts.slice(1).join(",") || "").trim(); // dozvoli zarez u value
      if (field) {
        newRows.push({ id: id++, field, value });
      }
    }

    if (newRows.length) {
      rows = newRows;
      selectedRowId = rows[0] ? rows[0].id : null;
      renderRows();
      await saveToDocument();
      showStatus(`✓ Importovano ${newRows.length} polja iz CSV`, true);
    }
  };

  input.click();
}

/* =========================
   WORD ACTIONS
   ========================= */

async function insertFieldWithMeta() {
  const key = getSelectedKey();
  if (!key) {
    showStatus("⚠ UBACI POLJE: prvo klikni red u tabeli", true);
    return;
  }

  const pick = await openInsertModal(key);
  if (!pick) return;

  const tag = makeTag(key, pick.type, pick.fmt);

  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const cc = range.insertContentControl();
    cc.tag = tag;
    cc.title = key;
    cc.appearance = "BoundingBox";
    cc.insertText(`{${key}}`, "Replace");
    await context.sync();
  });

  await saveToDocument();
  showStatus(`✓ Ubačeno polje: {${key}}`, true);
}

async function fillAll() {
  const map = getMap();
  const askedToday = new Set();

  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let changed = 0;
    let skipped = 0;
    let usedToday = 0;

    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) { skipped++; continue; }

      let raw = map.get(meta.key);

      if (meta.type === "D" && (!raw || raw.trim() === "")) {
        if (!askedToday.has(meta.key)) {
          askedToday.add(meta.key);
          const yes = await confirmToday(meta.key);
          if (yes) {
            raw = todayRawDDMMYYYY();
            usedToday++;

            rows = rows.map((r) => {
              if (normalizeKey(r.field) === meta.key) return { ...r, value: raw };
              return r;
            });
            renderRows();
          }
        }
      }

      if (raw === undefined) { skipped++; continue; }
      if (meta.type !== "D" && raw.trim() === "") { skipped++; continue; }
      if (meta.type === "D" && raw.trim() === "") { skipped++; continue; }

      const out = applyFormat(raw, meta.type, meta.fmt);
      cc.insertText(out, "Replace");
      changed++;
    }

    await context.sync();

    let msg = `✓ Popunjeno ${changed} polja`;
    if (usedToday > 0) msg += `, današnji datum ${usedToday}x`;
    if (skipped > 0) msg += `, preskočeno ${skipped}`;
    showStatus(msg, true);
  });

  await saveToDocument();
}

async function clearFieldsToPlaceholder() {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let cleared = 0;
    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;
      cc.insertText(`{${meta.key}}`, "Replace");
      cleared++;
    }

    await context.sync();
    showStatus(`✓ Očišćeno ${cleared} polja`, true);
  });

  await saveToDocument();
}

async function deleteControlsKeepText() {
  const confirm = window.confirm(
    "⚠ UPOZORENJE\n\n" +
    "Ova akcija će:\n" +
    "• Obrisati sve kontrole polja\n" +
    "• Zadržati tekst u dokumentu\n" +
    "• Obrisati sačuvane podatke (XML)\n\n" +
    "Da li ste sigurni?"
  );

  if (!confirm) return;

  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let removed = 0;
    for (let i = ccs.items.length - 1; i >= 0; i--) {
      const cc = ccs.items[i];
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;
      cc.delete(true); // keepContent = true
      removed++;
    }

    await context.sync();
    showStatus(`✓ Uklonjeno ${removed} kontrola (tekst ostao)`, true);
  });

  try {
    await Word.run(async (context) => {
      const parts = context.document.customXmlParts.getByNamespace(XML_NS);
      parts.load("items");
      await context.sync();
      parts.items.forEach((p) => p.delete());
      await context.sync();
    });
  } catch (e) {
    console.error("XML delete failed:", e);
  }
}

/* =========================
   KEYBOARD SHORTCUTS
   ========================= */

function setupKeyboardShortcuts() {
  document.addEventListener("keydown", (e) => {
    const activeEl = document.activeElement;
    const isInput = activeEl && (activeEl.tagName === "INPUT" || activeEl.tagName === "TEXTAREA");

    if (e.ctrlKey && e.key === "Enter") {
      e.preventDefault();
      insertFieldWithMeta();
      return;
    }

    if (e.ctrlKey && (e.key === "s" || e.key === "S")) {
      e.preventDefault();
      saveToDocument().then(() => showStatus("✓ Sačuvano", true));
      return;
    }

    if (e.key === "Delete" && !isInput && selectedRowId) {
      e.preventDefault();
      deleteRow(selectedRowId);
      return;
    }

    if (e.ctrlKey && (e.key === "e" || e.key === "E")) {
      e.preventDefault();
      exportCSV();
      return;
    }

    if (e.ctrlKey && (e.key === "i" || e.key === "I")) {
      e.preventDefault();
      importCSV();
      return;
    }
  });
}

/* =========================
   UI WIRING
   ========================= */

function setActive(btnId) {
  const ids = ["btnInsert", "btnFill", "btnClear", "btnDelete"];
  ids.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.classList.remove("active");
  });
  const b = document.getElementById(btnId);
  if (b) b.classList.add("active");
}

function wireUI() {
  const btnAdd = $("btnAddRow");
  if (btnAdd) btnAdd.addEventListener("click", addRow);

  const btnInsert = $("btnInsert");
  if (btnInsert) {
    btnInsert.addEventListener("click", async () => {
      setActive("btnInsert");
      showStatus("");
      try { await insertFieldWithMeta(); }
      catch (e) { showStatus(`⚠ Greška: ${e && e.message ? e.message : String(e)}`, true); }
    });
  }

  const btnFill = $("btnFill");
  if (btnFill) {
    btnFill.addEventListener("click", async () => {
      setActive("btnFill");
      showStatus("");
      try { await fillAll(); }
      catch (e) { showStatus(`⚠ Greška: ${e && e.message ? e.message : String(e)}`, true); }
    });
  }

  const btnClear = $("btnClear");
  if (btnClear) {
    btnClear.addEventListener("click", async () => {
      setActive("btnClear");
      showStatus("");
      try { await clearFieldsToPlaceholder(); }
      catch (e) { showStatus(`⚠ Greška: ${e && e.message ? e.message : String(e)}`, true); }
    });
  }

  const btnDelete = $("btnDelete");
  if (btnDelete) {
    btnDelete.addEventListener("click", async () => {
      setActive("btnDelete");
      showStatus("");
      try { await deleteControlsKeepText(); }
      catch (e) { showStatus(`⚠ Greška: ${e && e.message ? e.message : String(e)}`, true); }
    });
  }

  const btnExport = $("btnExportCSV");
  if (btnExport) {
    btnExport.addEventListener("click", () => {
      try { exportCSV(); }
      catch (e) { showStatus(`⚠ Greška: ${e && e.message ? e.message : String(e)}`, true); }
    });
  }

  const btnImport = $("btnImportCSV");
  if (btnImport) {
    btnImport.addEventListener("click", () => {
      try { importCSV(); }
      catch (e) { showStatus(`⚠ Greška: ${e && e.message ? e.message : String(e)}`, true); }
    });
  }

  if (selectedRowId == null && rows.length) selectedRowId = rows[0].id;
  renderRows();
  setupKeyboardShortcuts();
}

Office.onReady(async () => {
  wireUI();

  try {
    await loadFromDocument();
  } catch (e) {
    console.error("Initial load failed:", e);
  }
});
