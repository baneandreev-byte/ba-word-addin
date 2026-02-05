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
  return document.getElementById(id) || null;
}

function setStatus(msg, kind = "info") {
  const s = elMaybe("status");
  if (!s) return;
  s.textContent = msg;
  s.className = `status ${kind}`;
  // auto hide after a bit
  clearTimeout(setStatus._t);
  setStatus._t = setTimeout(() => {
    s.textContent = "";
    s.className = "status";
  }, 3500);
}

// ---------- row helpers ----------
function ensureDefaults(r) {
  return {
    id: r.id,
    field: r.field ?? "",
    value: r.value ?? "",
    type: r.type ?? "text",
    format: r.format ?? "text:auto",
  };
}

function normalizeKey(s) {
  return String(s ?? "").trim();
}

function token(key) {
  return `{${key}}`;
}

// ---------- tag format in content control ----------
// Tag format: BA_FIELD|key=<KEY>|type=<TYPE>|format=<FORMAT>
function makeTag(key, type, format) {
  const k = normalizeKey(key);
  const t = (type || "text").trim();
  const f = (format || "text:auto").trim();
  return `BA_FIELD|key=${k}|type=${t}|format=${f}`;
}

function parseTag(tag) {
  const s = String(tag || "");
  if (!s.startsWith("BA_FIELD|")) return null;
  const parts = s.split("|").slice(1);
  const out = {};
  for (const p of parts) {
    const [k, ...rest] = p.split("=");
    out[k] = rest.join("=");
  }
  if (!out.key) return null;
  return {
    key: out.key,
    type: out.type || "text",
    format: out.format || "text:auto",
  };
}

// ---------- formatting ----------
function applyFormat(type, format, rawValue) {
  const v = String(rawValue ?? "");

  if (!v) return "";

  // minimal formatting rules
  if (type === "number") {
    const n = Number(String(v).replace(",", "."));
    if (Number.isNaN(n)) return v;
    if (format === "number:int") return String(Math.round(n));
    if (format === "number:2") return n.toFixed(2);
    return String(n);
  }

  if (type === "date") {
    // expect yyyy-mm-dd or dd.mm.yyyy; keep as-is if unknown
    if (format === "date:today") {
      const d = new Date();
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const yyyy = d.getFullYear();
      return `${dd}.${mm}.${yyyy}`;
    }
    return v;
  }

  // text
  return v;
}

function buildValueMap() {
  const map = new Map();
  for (const r0 of rows) {
    const r = ensureDefaults(r0);
    const key = normalizeKey(r.field);
    if (!key) continue;
    const raw = (r.value ?? "").trim();
    const formatted = applyFormat(r.type, r.format, raw);
    map.set(key, { raw, formatted });
  }
  return map;
}

// ---------- render ----------
function renderRows() {
  const tbodyLeft = elMaybe("tbody-left");
  const tbodyRight = elMaybe("tbody-right");
  if (!tbodyLeft || !tbodyRight) return;

  tbodyLeft.innerHTML = "";
  tbodyRight.innerHTML = "";

  for (const r0 of rows) {
    const r = ensureDefaults(r0);

    // LEFT table: fields list
    const trL = document.createElement("tr");
    trL.dataset.id = String(r.id);
    if (r.id === selectedRowId) trL.classList.add("selected");

    const tdField = document.createElement("td");
    tdField.textContent = r.field || "";
    trL.appendChild(tdField);

    trL.addEventListener("click", () => {
      selectedRowId = r.id;
      renderRows();
      syncEditorFromSelected();
    });

    tbodyLeft.appendChild(trL);

    // RIGHT table: values
    const trR = document.createElement("tr");
    trR.dataset.id = String(r.id);
    if (r.id === selectedRowId) trR.classList.add("selected");

    const tdValue = document.createElement("td");
    tdValue.textContent = r.value || "";
    trR.appendChild(tdValue);

    trR.addEventListener("click", () => {
      selectedRowId = r.id;
      renderRows();
      syncEditorFromSelected();
    });

    tbodyRight.appendChild(trR);
  }
}

function getSelectedRow() {
  return rows.find((r) => r.id === selectedRowId) || rows[0] || null;
}

function syncEditorFromSelected() {
  const r = getSelectedRow();
  if (!r) return;

  const field = elMaybe("inp-field");
  const value = elMaybe("inp-value");
  const type = elMaybe("sel-type");
  const format = elMaybe("sel-format");
  if (field) field.value = r.field || "";
  if (value) value.value = r.value || "";
  if (type) type.value = r.type || "text";
  if (format) format.value = r.format || "text:auto";
}

function syncSelectedFromEditor() {
  const r = getSelectedRow();
  if (!r) return;

  const field = elMaybe("inp-field");
  const value = elMaybe("inp-value");
  const type = elMaybe("sel-type");
  const format = elMaybe("sel-format");

  r.field = field ? field.value : r.field;
  r.value = value ? value.value : r.value;
  r.type = type ? type.value : r.type;
  r.format = format ? format.value : r.format;
}

// ---------- persistence (CustomXml) ----------
const XML_NS = "http://biroa.local/ba-word-addin";
const XML_ROOT = "BAWordAddinState";

function buildStateXml() {
  const esc = (s) =>
    String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");

  const items = rows
    .map(ensureDefaults)
    .filter((r) => normalizeKey(r.field))
    .map(
      (r) =>
        `<item field="${esc(r.field)}" value="${esc(r.value)}" type="${esc(
          r.type
        )}" format="${esc(r.format)}" />`
    )
    .join("");

  return `<?xml version="1.0" encoding="UTF-8"?>
<${XML_ROOT} xmlns="${XML_NS}">
  ${items}
</${XML_ROOT}>`;
}

async function saveStateToDocument() {
  const xml = buildStateXml();

  await Word.run(async (context) => {
    const parts = context.document.customXmlParts;
    parts.load("items");
    await context.sync();

    // delete existing BA state
    for (const p of parts.items) {
      p.load("namespaceUri");
    }
    await context.sync();

    for (const p of parts.items) {
      if (p.namespaceUri === XML_NS) p.delete();
    }
    await context.sync();

    context.document.customXmlParts.add(xml);
    await context.sync();
  });
}

async function loadStateFromDocument() {
  await Word.run(async (context) => {
    const parts = context.document.customXmlParts;
    parts.load("items");
    await context.sync();

    for (const p of parts.items) {
      p.load("namespaceUri");
    }
    await context.sync();

    const mine = parts.items.find((p) => p.namespaceUri === XML_NS);
    if (!mine) return;

    const xml = mine.getXml();
    await context.sync();

    const str = xml.value || "";
    // simple parse: extract <item ... />
    const items = [];
    const re = /<item\s+([^/>]+?)\s*\/>/g;
    let m;
    while ((m = re.exec(str))) {
      const attrs = m[1];
      const get = (name) => {
        const rm = new RegExp(`${name}="([^"]*)"`);
        const mm = rm.exec(attrs);
        if (!mm) return "";
        return mm[1]
          .replace(/&quot;/g, '"')
          .replace(/&apos;/g, "'")
          .replace(/&gt;/g, ">")
          .replace(/&lt;/g, "<")
          .replace(/&amp;/g, "&");
      };
      items.push({
        field: get("field"),
        value: get("value"),
        type: get("type") || "text",
        format: get("format") || "text:auto",
      });
    }

    if (items.length) {
      rows = items.map((it, idx) => ({
        id: idx + 1,
        field: it.field,
        value: it.value,
        type: it.type,
        format: it.format,
      }));
      if (!rows.length)
        rows = [{ id: 1, field: "", value: "", type: "text", format: "text:auto" }];
      selectedRowId = rows[0].id;
    }
  });
}

async function deleteSavedStateFromDocument() {
  await Word.run(async (context) => {
    const parts = context.document.customXmlParts;
    parts.load("items");
    await context.sync();

    for (const p of parts.items) {
      p.load("namespaceUri");
    }
    await context.sync();

    for (const p of parts.items) {
      if (p.namespaceUri === XML_NS) p.delete();
    }
    await context.sync();
  });
}

// ---------- Word operations ----------
async function insertFieldAtSelection() {
  syncSelectedFromEditor();
  const r = getSelectedRow();
  if (!r) return;

  const key = normalizeKey(r.field);
  if (!key) {
    setStatus("Unesi naziv polja.", "warn");
    return;
  }

  await Word.run(async (context) => {
    const sel = context.document.getSelection();
    const cc = sel.insertContentControl();
    cc.tag = makeTag(key, r.type, r.format);
    cc.title = key;
    cc.appearance = "Tags";
    cc.insertText(token(key), Word.InsertLocation.replace);
    await context.sync();
    setStatus(`Dodato polje: ${key}`, "info");
  });

  await saveStateToDocument();
}

async function scanFieldsFromDocument() {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag,title");
    await context.sync();

    const found = [];
    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;
      found.push({
        field: meta.key,
        value: "",
        type: meta.type || "text",
        format: meta.format || "text:auto",
      });
    }

    const uniq = new Map();
    for (const f of found) {
      const k = normalizeKey(f.field);
      if (!k) continue;
      if (!uniq.has(k))
        uniq.set(k, { field: k, value: "", type: f.type, format: f.format });
    }

    const arr = Array.from(uniq.values());
    rows = arr.length
      ? arr.map((it, idx) => ({ id: idx + 1, ...it }))
      : [{ id: 1, field: "", value: "", type: "text", format: "text:auto" }];

    selectedRowId = rows[0].id;
  });

  renderRows();
  syncEditorFromSelected();
  await saveStateToDocument();
  setStatus("Skenirano.", "info");
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

      // VAŽNO: upis u sadržaj kontrole (ne preko getRange()), da polje ostane
      cc.insertText(out, Word.InsertLocation.replace);

      filled++;
    }
    await context.sync();
    setStatus(`Popunjeno ${filled} polja.`, "info");
  });

  await saveStateToDocument();
}

async function clearFieldsKeepControls() {
  // UI: obriši vrednosti u desnoj tabeli
  rows = rows.map((r) => ({ ...r, value: "" }));
  renderRows();

  // Word: vrati sva BA polja na {KEY}, ali ostavi kontrole
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let cleared = 0;
    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;

      // VAŽNO: upis u sadržaj kontrole, da kontrola ostane
      cc.insertText(token(meta.key), Word.InsertLocation.replace);

      cleared++;
    }
    await context.sync();
    setStatus(`Očišćeno ${cleared} polja i obrisane vrednosti.`, "info");
  });

  // Persist: sačuvaj prazne vrednosti
  await saveStateToDocument();
}

async function deleteControlsLeaveTextAndXml() {
  // cilj: ukloniti SVE tragove plugina (kontrole + XML), i ostaviti samo običan tekst
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

      // bez token fallback-a: ako nema vrednosti -> prazno
      const out = map.get(meta.key)?.formatted ?? "";

      // upiši tekst u kontrolu, pa obriši kontrolu
      cc.insertText(out, Word.InsertLocation.replace);
      cc.delete(false);
      removed++;
    }
    await context.sync();
    setStatus(`Obrisane kontrole: ${removed}.`, "info");
  });

  // ukloni persistence (CustomXml)
  await deleteSavedStateFromDocument();

  // očisti UI (opciono, ali praktično)
  rows = [{ id: 1, field: "", value: "", type: "text", format: "text:auto" }];
  selectedRowId = rows[0]?.id ?? null;
  renderRows();

  setStatus("Dokument je očišćen od plugina (bez kontrola i XML).", "info");
}

// ---------- CSV ----------
function csvEscapeCell(s, delimiter = ";") {
  const v = String(s ?? "");
  const mustQuote =
    v.includes(delimiter) || v.includes('"') || v.includes("\n") || v.includes("\r");
  if (!mustQuote) return v;
  return `"${v.replace(/"/g, '""')}"`;
}

function exportCSV() {
  const delimiter = ";";
  const lines = [];
  for (const r of rows) {
    const f = (r.field || "").trim();
    const v = (r.value || "").trim();
    if (!f) continue;
    lines.push(`${csvEscapeCell(f, delimiter)}${delimiter}${csvEscapeCell(v, delimiter)}`);
  }

  const csv = lines.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "biroa-fields.csv";
  a.click();
  URL.revokeObjectURL(url);

  setStatus(`Eksportovano ${lines.length} polja u CSV (delimiter ';').`, "info");
}

function parseCsvLine(line, delimiter = ";") {
  const out = [];
  let cur = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];

    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cur += ch;
      }
    } else {
      if (ch === '"') inQuotes = true;
      else if (ch === delimiter) {
        out.push(cur);
        cur = "";
      } else {
        cur += ch;
      }
    }
  }
  out.push(cur);
  return out;
}

async function importCSV() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".csv";

  input.onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const text = await file.text();
    const lines = text
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);

    const newRows = [];
    let id = 1;

    for (const line of lines) {
      // novi standard ';', fallback za stare fajlove ','
      let parts = parseCsvLine(line, ";");
      let usedDelim = ";";
      if (parts.length < 2) {
        parts = parseCsvLine(line, ",");
        usedDelim = ",";
      }

      const field = (parts[0] || "").trim();
      const value = (parts.slice(1).join(usedDelim) || "").trim();
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

// ---------- wiring ----------
function bindUi() {
  const btnAdd = elMaybe("btn-add");
  const btnScan = elMaybe("btn-scan");
  const btnFill = elMaybe("btn-fill");
  const btnClear = elMaybe("btn-clear");
  const btnDelete = elMaybe("btn-delete");
  const btnExport = elMaybe("btn-export");
  const btnImport = elMaybe("btn-import");

  if (btnAdd) btnAdd.addEventListener("click", insertFieldAtSelection);
  if (btnScan) btnScan.addEventListener("click", scanFieldsFromDocument);
  if (btnFill) btnFill.addEventListener("click", fillFieldsFromTable);
  if (btnClear) btnClear.addEventListener("click", clearFieldsKeepControls);
  if (btnDelete) btnDelete.addEventListener("click", deleteControlsLeaveTextAndXml);
  if (btnExport) btnExport.addEventListener("click", exportCSV);
  if (btnImport) btnImport.addEventListener("click", importCSV);

  const field = elMaybe("inp-field");
  const value = elMaybe("inp-value");
  const type = elMaybe("sel-type");
  const format = elMaybe("sel-format");

  const onChange = async () => {
    syncSelectedFromEditor();
    renderRows();
    await saveStateToDocument();
  };

  if (field) field.addEventListener("input", onChange);
  if (value) value.addEventListener("input", onChange);
  if (type) type.addEventListener("change", onChange);
  if (format) format.addEventListener("change", onChange);
}

Office.onReady(async () => {
  try {
    await loadStateFromDocument();
  } catch (e) {
    // ignore
  }

  renderRows();
  syncEditorFromSelected();
  bindUi();
});
