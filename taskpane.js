/* global Office, Word */

let rows = [
  { id: 1, field: "", value: "", type: "text", format: "text:auto" },
  { id: 2, field: "", value: "", type: "text", format: "text:auto" },
  { id: 3, field: "", value: "", type: "text", format: "text:auto" },
  { id: 4, field: "", value: "", type: "text", format: "text:auto" },
];

let selectedRowId = null;
let currentModal = { rowId: null, fieldName: "" };

// Format definitions
const FORMATS = {
  text: [
    { value: "text:auto", label: "Automatski", hint: "Bez formatiranja" },
    { value: "text:upper", label: "VELIKA SLOVA", hint: "Sve veliko" },
    { value: "text:lower", label: "mala slova", hint: "Sve malo" },
    { value: "text:title", label: "Naslov", hint: "Prvo Veliko" },
  ],
  date: [
    { value: "date:today", label: "Danas", hint: "Automatski današnji datum" },
    { value: "date:dd.mm.yyyy", label: "dd.mm.yyyy", hint: "21.12.2024" },
    { value: "date:yyyy-mm-dd", label: "yyyy-mm-dd", hint: "2024-12-21" },
  ],
  number: [
    { value: "number:auto", label: "Automatski", hint: "Kao što je upisano" },
    { value: "number:int", label: "Ceo broj", hint: "Bez decimala" },
    { value: "number:2", label: "2 decimale", hint: "123.45" },
  ],
};

// ---------- DOM helpers ----------
function el(id) {
  return document.getElementById(id);
}

function elMaybe(id) {
  return document.getElementById(id) || null;
}

function showMessage(msg, type = "info") {
  const s = elMaybe("status");
  if (!s) return;
  s.textContent = msg;
  s.style.display = "block";
  s.className = `status ${type}`;
  setTimeout(() => {
    s.style.display = "none";
  }, 3000);
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

  if (type === "number") {
    const n = Number(String(v).replace(",", "."));
    if (Number.isNaN(n)) return v;
    if (format === "number:int") return String(Math.round(n));
    if (format === "number:2") return n.toFixed(2);
    return String(n);
  }

  if (type === "date") {
    if (format === "date:today") {
      const d = new Date();
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const yyyy = d.getFullYear();
      return `${dd}.${mm}.${yyyy}`;
    }
    return v;
  }

  // text formatting
  if (type === "text") {
    if (format === "text:upper") return v.toUpperCase();
    if (format === "text:lower") return v.toLowerCase();
    if (format === "text:title") {
      return v.replace(/\b\w/g, (c) => c.toUpperCase());
    }
  }

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
  const container = elMaybe("rows");
  if (!container) return;

  container.innerHTML = "";

  for (const r0 of rows) {
    const r = ensureDefaults(r0);

    const rowDiv = document.createElement("div");
    rowDiv.className = "row";
    rowDiv.dataset.id = String(r.id);

    // Field cell
    const cellField = document.createElement("div");
    cellField.className = "cell";
    const inputField = document.createElement("input");
    inputField.type = "text";
    inputField.placeholder = "Naziv polja";
    inputField.value = r.field || "";
    inputField.addEventListener("input", (e) => {
      r.field = e.target.value;
      saveStateToDocument();
    });
    inputField.addEventListener("click", () => {
      selectedRowId = r.id;
      openModal(r.id, r.field || "");
    });
    cellField.appendChild(inputField);

    // Value cell
    const cellValue = document.createElement("div");
    cellValue.className = "cell";
    const inputValue = document.createElement("input");
    inputValue.type = "text";
    inputValue.placeholder = "Vrednost";
    inputValue.value = r.value || "";
    inputValue.addEventListener("input", (e) => {
      r.value = e.target.value;
      saveStateToDocument();
    });
    cellValue.appendChild(inputValue);

    // Delete button cell
    const cellDel = document.createElement("div");
    cellDel.className = "del";
    const btnDel = document.createElement("button");
    btnDel.textContent = "×";
    btnDel.addEventListener("click", () => deleteRow(r.id));
    cellDel.appendChild(btnDel);

    rowDiv.appendChild(cellField);
    rowDiv.appendChild(cellValue);
    rowDiv.appendChild(cellDel);

    container.appendChild(rowDiv);
  }
}

function deleteRow(rowId) {
  if (rows.length <= 1) {
    showMessage("Ne možete obrisati poslednji red!", "error");
    return;
  }
  rows = rows.filter((r) => r.id !== rowId);
  renderRows();
  saveStateToDocument();
}

function addRow() {
  const newId = rows.length > 0 ? Math.max(...rows.map((r) => r.id)) + 1 : 1;
  rows.push({
    id: newId,
    field: "",
    value: "",
    type: "text",
    format: "text:auto",
  });
  renderRows();
  saveStateToDocument();
}

// ---------- Modal ----------
function openModal(rowId, fieldName) {
  currentModal = { rowId, fieldName };

  const modal = el("modal");
  const backdrop = el("modalBackdrop");
  const fieldNameSpan = el("modalFieldName");
  const formatSelect = el("formatSelect");

  fieldNameSpan.textContent = fieldName || "(prazno)";

  // Get current row data
  const row = rows.find((r) => r.id === rowId);
  if (!row) return;

  // Set type radio
  const radios = document.querySelectorAll('input[name="ftype"]');
  radios.forEach((radio) => {
    radio.checked = radio.value === row.type;
  });

  // Populate format dropdown
  updateFormatSelect(row.type, row.format);

  // Show modal
  modal.classList.remove("hidden");
  backdrop.classList.remove("hidden");
}

function closeModal() {
  const modal = el("modal");
  const backdrop = el("modalBackdrop");
  modal.classList.add("hidden");
  backdrop.classList.add("hidden");
}

function updateFormatSelect(type, currentFormat) {
  const formatSelect = el("formatSelect");
  const formatHint = el("formatHint");

  formatSelect.innerHTML = "";

  const formats = FORMATS[type] || FORMATS.text;
  formats.forEach((fmt) => {
    const option = document.createElement("option");
    option.value = fmt.value;
    option.textContent = fmt.label;
    if (fmt.value === currentFormat) {
      option.selected = true;
    }
    formatSelect.appendChild(option);
  });

  // Update hint
  const selectedFmt = formats.find((f) => f.value === formatSelect.value);
  formatHint.textContent = selectedFmt ? selectedFmt.hint : "";

  // Update hint on change
  formatSelect.addEventListener("change", () => {
    const fmt = formats.find((f) => f.value === formatSelect.value);
    formatHint.textContent = fmt ? fmt.hint : "";
  });
}

async function handleModalOk() {
  const row = rows.find((r) => r.id === currentModal.rowId);
  if (!row) return;

  // Get selected type
  const selectedType = document.querySelector('input[name="ftype"]:checked')?.value || "text";
  const selectedFormat = el("formatSelect").value;

  row.type = selectedType;
  row.format = selectedFormat;

  await saveStateToDocument();
  closeModal();

  // Now insert the field
  await insertFieldAtCursor(row);
}

async function insertFieldAtCursor(row) {
  if (!row.field) {
    showMessage("Naziv polja je prazan!", "error");
    return;
  }

  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      const cc = range.insertContentControl();
      cc.title = row.field;
      cc.tag = makeTag(row.field, row.type, row.format);
      cc.insertText(token(row.field), Word.InsertLocation.replace);
      cc.appearance = "Tags";
      cc.color = "#3b82f6";

      await context.sync();
      showMessage(`Polje "${row.field}" ubačeno!`, "success");
    });
  } catch (error) {
    console.error("Error inserting field:", error);
    showMessage("Greška: " + error.message, "error");
  }
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
  try {
    const xml = buildStateXml();

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

      context.document.customXmlParts.add(xml);
      await context.sync();
    });
  } catch (error) {
    console.error("Error saving state:", error);
  }
}

async function loadStateFromDocument() {
  try {
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
        rows = items.map((it, idx) => ({ id: idx + 1, ...it }));
      }
    });
  } catch (error) {
    console.error("Error loading state:", error);
  }
}

async function deleteSavedStateFromDocument() {
  try {
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
  } catch (error) {
    console.error("Error deleting state:", error);
  }
}

// ---------- Main Actions ----------
async function fillFields() {
  const map = buildValueMap();

  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag");
      await context.sync();

      let filled = 0;
      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta) continue;
        const out = map.get(meta.key)?.formatted ?? token(meta.key);

        // KRITIČNO: koristimo cc.text property da postavimo vrednost
        // ali OSTAVLJAMO content control živ i editabilan
        cc.text = out;

        filled++;
      }
      await context.sync();
      showMessage(`Popunjeno ${filled} polja.`, "success");
    });

    await saveStateToDocument();
  } catch (error) {
    console.error("Error filling fields:", error);
    showMessage("Greška: " + error.message, "error");
  }
}

async function clearFields() {
  // Očisti vrednosti u UI
  rows = rows.map((r) => ({ ...r, value: "" }));
  renderRows();

  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag");
      await context.sync();

      let cleared = 0;
      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta) continue;

        // Vrati na {POLJE} placeholder ali ostavi content control
        cc.text = token(meta.key);

        cleared++;
      }
      await context.sync();
      showMessage(`Očišćeno ${cleared} polja.`, "success");
    });

    await saveStateToDocument();
  } catch (error) {
    console.error("Error clearing fields:", error);
    showMessage("Greška: " + error.message, "error");
  }
}

async function deleteAllControls() {
  const map = buildValueMap();

  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag");
      await context.sync();

      let removed = 0;
      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta) continue;

        // Postavi tekst (ili prazno ako nema vrednosti)
        const out = map.get(meta.key)?.formatted ?? "";
        cc.text = out;

        // OBRIŠI content control ali OSTAVI tekst
        cc.delete(false); // false = keep text

        removed++;
      }
      await context.sync();
      showMessage(`Obrisano ${removed} kontrola. Tekst je ostao.`, "success");
    });

    // Obriši i XML persistence
    await deleteSavedStateFromDocument();

    // Reset UI
    rows = [{ id: 1, field: "", value: "", type: "text", format: "text:auto" }];
    renderRows();
  } catch (error) {
    console.error("Error deleting controls:", error);
    showMessage("Greška: " + error.message, "error");
  }
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

  showMessage(`Eksportovano ${lines.length} polja.`, "success");
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
      renderRows();
      await saveStateToDocument();
      showMessage(`Importovano ${newRows.length} polja.`, "success");
    }
  };

  input.click();
}

// ---------- UI Binding ----------
function bindUI() {
  // Tab buttons
  el("btnInsert")?.addEventListener("click", () => {
    showMessage("Klikni red u tabeli, pa klikni u dokument gde želiš polje.", "info");
  });
  el("btnFill")?.addEventListener("click", fillFields);
  el("btnClear")?.addEventListener("click", clearFields);
  el("btnDelete")?.addEventListener("click", deleteAllControls);

  // Action buttons
  el("btnAddRow")?.addEventListener("click", addRow);
  el("btnExportCSV")?.addEventListener("click", exportCSV);
  el("btnImportCSV")?.addEventListener("click", importCSV);

  // Modal buttons
  el("btnModalClose")?.addEventListener("click", closeModal);
  el("btnModalCancel")?.addEventListener("click", closeModal);
  el("btnModalOk")?.addEventListener("click", handleModalOk);
  el("modalBackdrop")?.addEventListener("click", closeModal);

  // Type radio change
  document.querySelectorAll('input[name="ftype"]').forEach((radio) => {
    radio.addEventListener("change", (e) => {
      const row = rows.find((r) => r.id === currentModal.rowId);
      if (row) {
        updateFormatSelect(e.target.value, row.format);
      }
    });
  });
}

// ---------- Init ----------
Office.onReady(async () => {
  try {
    await loadStateFromDocument();
  } catch (e) {
    console.error("Load state error:", e);
  }

  renderRows();
  bindUI();
  showMessage("Add-in spreman!", "success");
});
