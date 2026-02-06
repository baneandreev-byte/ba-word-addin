/* global Office, Word */

let rows = [
  { id: 1, field: "", value: "", type: "text", format: "text:auto" },
  { id: 2, field: "", value: "", type: "text", format: "text:auto" },
  { id: 3, field: "", value: "", type: "text", format: "text:auto" },
  { id: 4, field: "", value: "", type: "text", format: "text:auto" },
];

let selectedRowId = rows[0]?.id ?? null;

// ---------- DOM helpers ----------
function el(id) {
  return document.getElementById(id);
}

function elMaybe(id) {
  return document.getElementById(id);
}

function show(id, on) {
  const n = elMaybe(id);
  if (!n) return;
  n.classList.toggle("hidden", !on);
}

function setStatus(msg, kind = "info") {
  const s = elMaybe("status");
  if (!s) return;
  if (!msg) {
    s.style.display = "none";
    s.textContent = "";
    return;
  }
  s.style.display = "block";
  s.textContent = msg;
  s.style.borderColor = kind === "error" ? "#fca5a5" : "#93c5fd";
  s.style.background = kind === "error" ? "#fef2f2" : "#eff6ff";
  s.style.color = "#111827";
}

function normalizeKey(s) {
  return (s || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[{}]/g, "")
    .replace(/[^A-Za-z0-9_]/g, "_")
    .replace(/_+/g, "_");
}

function token(key) {
  return `{${key}}`;
}

function ensureDefaults(r) {
  const type = r.type ?? "text";
  let format = r.format ?? "";
  if (!format) {
    if (type === "date") format = "date:MMMM_yyyy_cap";
    else if (type === "number") format = "num:number";
    else format = "text:auto";
  }
  return { ...r, type, format };
}

function getSelectedRow() {
  if (selectedRowId == null) return null;
  return rows.find((r) => r.id === selectedRowId) ?? null;
}

// ---------- Render table ----------
function renderRows() {
  const container = elMaybe("rows");
  if (!container) return;

  container.innerHTML = "";

  rows.forEach((r0) => {
    const r = ensureDefaults(r0);

    const wrap = document.createElement("div");
    wrap.className = "row";
    wrap.style.cursor = "pointer";
    if (selectedRowId === r.id) wrap.style.background = "#eff6ff";

    wrap.addEventListener("click", () => {
      selectedRowId = r.id;
      renderRows();
    });

    const c1 = document.createElement("div");
    c1.className = "cell";
    const i1 = document.createElement("input");
    i1.value = r.field ?? "";
    i1.placeholder = "Unesite polje...";
    i1.addEventListener("click", (e) => e.stopPropagation());
    i1.addEventListener("focus", () => {
      if (selectedRowId !== r.id) {
        selectedRowId = r.id;
        renderRows();
      }
    });
    i1.addEventListener("input", (e) => {
      const v = e.target.value;
      rows = rows.map((x) => (x.id === r.id ? { ...x, field: v } : x));
    });
    c1.appendChild(i1);

    const c2 = document.createElement("div");
    c2.className = "cell";
    const i2 = document.createElement("input");
    i2.value = r.value ?? "";
    i2.placeholder = "Unesite odgovor...";
    i2.addEventListener("click", (e) => e.stopPropagation());
    i2.addEventListener("focus", () => {
      if (selectedRowId !== r.id) {
        selectedRowId = r.id;
        renderRows();
      }
    });
    i2.addEventListener("input", (e) => {
      const v = e.target.value;
      rows = rows.map((x) => (x.id === r.id ? { ...x, value: v } : x));
    });
    c2.appendChild(i2);

    const c3 = document.createElement("div");
    c3.className = "del";
    const bDel = document.createElement("button");
    bDel.textContent = "×";
    bDel.title = "Obriši red";
    bDel.addEventListener("click", (e) => {
      e.stopPropagation();
      if (rows.length <= 1) return;
      rows = rows.filter((x) => x.id !== r.id);
      if (selectedRowId === r.id) selectedRowId = rows[0]?.id ?? null;
      renderRows();
    });
    c3.appendChild(bDel);

    wrap.appendChild(c1);
    wrap.appendChild(c2);
    wrap.appendChild(c3);
    container.appendChild(wrap);
  });
}

function addRow() {
  const newId = Math.max(...rows.map((r) => r.id), 0) + 1;
  rows.push({ id: newId, field: "", value: "", type: "text", format: "text:auto" });
  selectedRowId = newId;
  renderRows();
}

// ---------- Modal logic ----------
const FORMATS = {
  text: [
    { value: "text:auto", label: "AUTO (kako je upisano)" },
    { value: "text:upper", label: "UPPER (VELIKA SLOVA)" },
    { value: "text:lower", label: "lower (mala slova)" },
    { value: "text:sentence", label: "Sentence (Prvo slovo veliko)" },
  ],
  date: [
    { value: "date:MMMM_yyyy_cap", label: "MMMM yyyy (April 2025)", hint: "Mesec kao tekst, veliko prvo slovo." },
    { value: "date:d_MMMM_yyyy", label: "d. MMMM yyyy (15. april 2025)" },
    { value: "date:dd_MM_yyyy", label: "dd.MM.yyyy (15.04.2025)" },
  ],
  number: [
    { value: "num:number", label: "BROJ (1.234,56)" },
    { value: "num:integer", label: "CELI BROJ (1.235)" },
    { value: "num:currency_RSD", label: "RSD (valuta)" },
    { value: "num:currency_EUR", label: "EUR (valuta)" },
  ],
};

let modalResolve = null;

function openModal(fieldKey, initType, initFormat) {
  const nameEl = elMaybe("modalFieldName");
  if (nameEl) nameEl.textContent = fieldKey;

  // set radio
  const radios = Array.from(document.querySelectorAll('input[name="ftype"]'));
  radios.forEach((r) => (r.checked = r.value === initType));
  fillFormatSelect(initType, initFormat);

  show("modalBackdrop", true);
  show("modal", true);
}

function closeModal() {
  show("modalBackdrop", false);
  show("modal", false);
}

function fillFormatSelect(type, selected) {
  const sel = elMaybe("formatSelect");
  const hint = elMaybe("formatHint");
  if (!sel) return;

  sel.innerHTML = "";

  const opts = FORMATS[type];
  opts.forEach((o) => {
    const opt = document.createElement("option");
    opt.value = o.value;
    opt.textContent = o.label;
    sel.appendChild(opt);
  });

  const exists = opts.some((o) => o.value === selected);
  sel.value = exists ? selected : opts[0].value;

  if (hint) {
    const current = opts.find((o) => o.value === sel.value);
    hint.textContent = current?.hint ?? "";
    sel.onchange = () => {
      const cur = opts.find((o) => o.value === sel.value);
      hint.textContent = cur?.hint ?? "";
    };
  }
}

function getSelectedTypeFromModal() {
  const radios = Array.from(document.querySelectorAll('input[name="ftype"]'));
  const hit = radios.find((r) => r.checked);
  const v = hit?.value || "text";
  return v === "date" || v === "number" || v === "text" ? v : "text";
}

function waitForModalChoice(fieldKey, initType, initFormat) {
  return new Promise((resolve) => {
    modalResolve = resolve;
    openModal(fieldKey, initType, initFormat);
  });
}

// ---------- Content controls tags ----------
function makeTag(key, type, format) {
  return `BA|${key}|${type}|${format || ""}`;
}

function parseTag(tag) {
  if (!tag || !tag.startsWith("BA|")) return null;
  const parts = tag.split("|");
  if (parts.length < 4) return null;
  return { key: parts[1] || "", type: parts[2] || "text", format: parts.slice(3).join("|") || "" };
}

// ---------- Formatting ----------
function capFirst(s) {
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}

function formatText(val, format) {
  const v = val ?? "";
  if (format === "text:upper") return v.toUpperCase();
  if (format === "text:lower") return v.toLowerCase();
  if (format === "text:sentence") return capFirst(v.toLowerCase());
  return v;
}

function parseDateLoose(input) {
  const s = (input ?? "").trim();
  if (!s) return null;

  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));

  const dm = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\.?$/);
  if (dm) return new Date(Number(dm[3]), Number(dm[2]) - 1, Number(dm[1]));

  const t = Date.parse(s);
  return Number.isNaN(t) ? null : new Date(t);
}

function formatDate(val, format) {
  const dt = parseDateLoose(val);
  if (!dt) return "";
  const locale = "sr-Latn-RS";

  if (format === "date:MMMM_yyyy_cap") {
    const month = new Intl.DateTimeFormat(locale, { month: "long" }).format(dt);
    const year = new Intl.DateTimeFormat(locale, { year: "numeric" }).format(dt);
    return `${capFirst(month)} ${year}`;
  }
  if (format === "date:d_MMMM_yyyy") {
    const day = new Intl.DateTimeFormat(locale, { day: "numeric" }).format(dt);
    const month = new Intl.DateTimeFormat(locale, { month: "long" }).format(dt);
    const year = new Intl.DateTimeFormat(locale, { year: "numeric" }).format(dt);
    return `${day}. ${month} ${year}`;
  }
  if (format === "date:dd_MM_yyyy") {
    const d = String(dt.getDate()).padStart(2, "0");
    const m = String(dt.getMonth() + 1).padStart(2, "0");
    const y = String(dt.getFullYear());
    return `${d}.${m}.${y}`;
  }

  const month = new Intl.DateTimeFormat(locale, { month: "long" }).format(dt);
  const year = new Intl.DateTimeFormat(locale, { year: "numeric" }).format(dt);
  return `${capFirst(month)} ${year}`;
}

function parseNumberLoose(input) {
  const s = (input ?? "").trim();
  if (!s) return null;
  const normalized = s.replace(/\s/g, "").replace(/\.(?=\d{3}\b)/g, "").replace(",", ".");
  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
}

function formatNumber(val, format) {
  const n = parseNumberLoose(val);
  if (n == null) return "";
  if (format === "num:currency_RSD") return new Intl.NumberFormat("sr-RS", { style: "currency", currency: "RSD" }).format(n);
  if (format === "num:currency_EUR") return new Intl.NumberFormat("sr-RS", { style: "currency", currency: "EUR" }).format(n);
  if (format === "num:integer") return new Intl.NumberFormat("sr-RS", { maximumFractionDigits: 0 }).format(n);
  const hasFrac = Math.abs(n % 1) > 0;
  return new Intl.NumberFormat("sr-RS", { minimumFractionDigits: hasFrac ? 2 : 0, maximumFractionDigits: hasFrac ? 2 : 0 }).format(n);
}

function applyFormat(type, format, raw) {
  if (type === "text") return formatText(raw, format);
  if (type === "date") return formatDate(raw, format);
  if (type === "number") return formatNumber(raw, format);
  return raw;
}

// ---------- Persistence (Custom XML) ----------
const XML_NS = "biroa/fields";
function escapeXml(s) {
  return (s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function saveStateToDocument() {
  await Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(XML_NS);
    parts.load("items");
    await context.sync();

    parts.items.forEach((p) => p.delete());
    await context.sync();

    let xml = `<BiroA xmlns="${XML_NS}">`;
    for (const r0 of rows) {
      const r = ensureDefaults(r0);
      const key = normalizeKey(r.field);
      if (!key) continue;
      xml += `<field name="${escapeXml(key)}" type="${escapeXml(r.type)}" format="${escapeXml(r.format)}"><value>${escapeXml(
        r.value ?? ""
      )}</value></field>`;
    }
    xml += `</BiroA>`;
    context.document.customXmlParts.add(xml);
    await context.sync();
  });
}

async function loadStateFromDocument() {
  await Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(XML_NS);
    parts.load("items");
    await context.sync();
    if (!parts.items.length) return;

    const xmlRes = parts.items[0].getXml();
    await context.sync();

    const xml = xmlRes.value || "";
    if (!xml) return;

    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const fields = Array.from(doc.getElementsByTagName("field"));

    const newRows = [];
    let idx = 1;
    for (const f of fields) {
      const name = f.getAttribute("name") || "";
      const type = f.getAttribute("type") || "text";
      const format = f.getAttribute("format") || "";
      const value = f.getElementsByTagName("value")[0]?.textContent || "";
      if (!name) continue;
      newRows.push({ id: idx++, field: name, value, type, format });
    }

    if (newRows.length) {
      rows = newRows;
      selectedRowId = rows[0].id;
      renderRows();
      setStatus(`Učitano ${rows.length} polja iz dokumenta.`, "info");
    }
  });
}

async function deleteSavedStateFromDocument() {
  await Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(XML_NS);
    parts.load("items");
    await context.sync();
    parts.items.forEach((p) => p.delete());
    await context.sync();
  });
}

// ---------- Word actions ----------
async function insertSelectedFieldViaModal() {
  const r0 = getSelectedRow();
  if (!r0) {
    setStatus("Izaberi red u tabeli pa klikni UBACI POLJE.", "error");
    return;
  }

  const key = normalizeKey(r0.field);
  if (!key) {
    setStatus("U izabranom redu unesi naziv polja (POLJE).", "error");
    return;
  }

  const base = ensureDefaults({ ...r0, field: key });

  // ako modal HTML ne postoji, prijavi jasno
  if (!elMaybe("modal") || !elMaybe("modalBackdrop") || !elMaybe("btnModalOk")) {
    setStatus("Modal UI nije prisutan u taskpane.html (nema modal elemenata).", "error");
    return;
  }

  const choice = await waitForModalChoice(key, base.type, base.format);
  if (!choice) {
    setStatus("Prekinuto.", "info");
    return;
  }

  rows = rows.map((x) => (x.id === base.id ? { ...x, field: key, type: choice.type, format: choice.format } : x));
  renderRows();

  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const cc = range.insertContentControl();
    cc.title = key;
    cc.tag = makeTag(key, choice.type, choice.format);
    cc.appearance = "BoundingBox";
    cc.placeholderText = token(key);
    cc.insertText(token(key), Word.InsertLocation.replace);
    await context.sync();
  });

  await saveStateToDocument();
  setStatus(`Ubačeno polje ${token(key)} (${choice.type}).`, "info");
}

function buildValueMap() {
  const map = new Map();
  for (const r0 of rows) {
    const r = ensureDefaults(r0);
    const key = normalizeKey(r.field);
    if (!key) continue;

    let raw = (r.value ?? "").trim();

    // date: if empty -> default to today (no dialogs)
    if (r.type === "date" && !raw) {
      const t = new Date();
      raw = `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
      rows = rows.map((x) => (x.id === r.id ? { ...x, value: raw } : x));
    }

    map.set(key, { formatted: applyFormat(r.type, r.format, raw) || token(key) });
  }
  return map;
}

async function fillFieldsFromTable() {
  const map = buildValueMap();
  renderRows();

  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let filled = 0;
    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;
      const out = map.get(meta.key)?.formatted ?? token(meta.key);
      cc.getRange().insertText(out, Word.InsertLocation.replace);
      filled++;
    }
    await context.sync();
    setStatus(`Popunjeno ${filled} polja.`, "info");
  });

  await saveStateToDocument();
}

async function clearFieldsKeepControls() {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let cleared = 0;
    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;
      cc.getRange().insertText(token(meta.key), Word.InsertLocation.replace);
      cleared++;
    }
    await context.sync();
    setStatus(`Očišćeno ${cleared} polja.`, "info");
  });

  await saveStateToDocument();
}

async function deleteControlsLeaveTextAndXml() {
  const map = buildValueMap();
  renderRows();

  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let removed = 0;
    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;
      const out = map.get(meta.key)?.formatted ?? "";
      cc.getRange().insertText(out, Word.InsertLocation.replace);
      cc.delete(false);
      removed++;
    }
    await context.sync();
    setStatus(`Obrisane kontrole: ${removed}.`, "info");
  });

  await deleteSavedStateFromDocument();
  setStatus("Obrisane kontrole i obrisani sačuvani podaci (XML).", "info");
}

// ---------- CSV Import/Export ----------
function exportCSV() {
  const lines = [];
  for (const r of rows) {
    const f = (r.field || "").trim();
    const v = (r.value || "").trim();
    if (!f) continue;
    lines.push(`${f},${v}`);
  }

  const csv = lines.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "biroa-fields.csv";
  a.click();
  URL.revokeObjectURL(url);

  setStatus(`Eksportovano ${lines.length} polja u CSV.`, "info");
}

async function importCSV() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".csv";

  input.onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const text = await file.text();
    const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);

    const newRows = [];
    let id = 1;

    for (const line of lines) {
      const parts = line.split(",");
      const field = (parts[0] || "").trim();
      const value = (parts.slice(1).join(",") || "").trim();
      if (field) {
        newRows.push({ id: id++, field, value, type: "text", format: "text:auto" });
      }
    }

    if (newRows.length) {
      rows = newRows;
      selectedRowId = rows[0]?.id ?? null;
      renderRows();
      await saveStateToDocument();
      setStatus(`Importovano ${newRows.length} polja iz CSV.`, "info");
    }
  };

  input.click();
}

// ---------- Wire UI ----------
function wireUI() {
  elMaybe("btnAddRow")?.addEventListener("click", () => addRow());

  elMaybe("btnInsert")?.addEventListener("click", async () => {
    setStatus("");
    try {
      await insertSelectedFieldViaModal();
    } catch (e) {
      setStatus(`Greška (UBACI POLJE): ${e?.message ?? String(e)}`, "error");
    }
  });

  elMaybe("btnFill")?.addEventListener("click", async () => {
    setStatus("");
    try {
      await fillFieldsFromTable();
    } catch (e) {
      setStatus(`Greška (POPUNI): ${e?.message ?? String(e)}`, "error");
    }
  });

  elMaybe("btnClear")?.addEventListener("click", async () => {
    setStatus("");
    try {
      await clearFieldsKeepControls();
    } catch (e) {
      setStatus(`Greška (OČISTI): ${e?.message ?? String(e)}`, "error");
    }
  });

  elMaybe("btnDelete")?.addEventListener("click", async () => {
    setStatus("");
    try {
      await deleteControlsLeaveTextAndXml();
    } catch (e) {
      setStatus(`Greška (OBRIŠI): ${e?.message ?? String(e)}`, "error");
    }
  });

  elMaybe("btnExportCSV")?.addEventListener("click", () => {
    try {
      exportCSV();
    } catch (e) {
      setStatus(`Greška (EXPORT): ${e?.message ?? String(e)}`, "error");
    }
  });

  elMaybe("btnImportCSV")?.addEventListener("click", () => {
    try {
      importCSV();
    } catch (e) {
      setStatus(`Greška (IMPORT): ${e?.message ?? String(e)}`, "error");
    }
  });

  // modal events (SAFE)
  elMaybe("btnModalClose")?.addEventListener("click", () => {
    closeModal();
    modalResolve?.(null);
    modalResolve = null;
  });
  elMaybe("btnModalCancel")?.addEventListener("click", () => {
    closeModal();
    modalResolve?.(null);
    modalResolve = null;
  });
  elMaybe("modalBackdrop")?.addEventListener("click", () => {
    closeModal();
    modalResolve?.(null);
    modalResolve = null;
  });

  // type change -> refill formats
  const radios = Array.from(document.querySelectorAll('input[name="ftype"]'));
  radios.forEach((r) =>
    r.addEventListener("change", () => {
      const t = getSelectedTypeFromModal();
      const defaults = FORMATS[t][0].value;
      fillFormatSelect(t, defaults);
    })
  );

  elMaybe("btnModalOk")?.addEventListener("click", () => {
    const t = getSelectedTypeFromModal();
    const sel = elMaybe("formatSelect");
    const f = sel ? sel.value : FORMATS[t][0].value;

    closeModal();
    modalResolve?.({ type: t, format: f });
    modalResolve = null;
  });
}

// ---------- Bootstrap ----------
Office.onReady(async () => {
  try {
    wireUI();
    renderRows();
    await loadStateFromDocument();
  } catch (e) {
    setStatus(`Greška pri startu: ${e?.message ?? String(e)}`, "error");
  }
});
