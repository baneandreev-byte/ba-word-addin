/* global Office, Word */

// ============================================
// VERZIJA: 2025-02-07 - V23
// ============================================
console.log("🔧 BA Word Add-in VERZIJA: 2025-02-07 - V23");
console.log("✅ Funkcije: Ubaci, Popuni (editabilno), Očisti (čuva vrednosti), Obriši (potvrda)");

let rows = [];
let selectedRowIndex = null;

// Format options per type
const FORMAT_OPTIONS = {
  text: [
    { value: "text:auto", label: "Automatski", hint: "" },
    { value: "text:upper", label: "VELIKA SLOVA", hint: "Primer: BEOGRAD" },
    { value: "text:lower", label: "mala slova", hint: "Primer: beograd" },
    { value: "text:title", label: "Naslov", hint: "Primer: Beograd" },
  ],
  date: [
    { value: "date:auto", label: "Kako je uneto", hint: "" },
    { value: "date:today", label: "Danas (dd.mm.yyyy)", hint: "Primer: 07.02.2025" },
    { value: "date:dd.mm.yyyy", label: "dd.mm.yyyy", hint: "Primer: 07.02.2025" },
    { value: "date:yyyy-mm-dd", label: "yyyy-mm-dd", hint: "Primer: 2025-02-07" },
    { value: "date:mmmm.yyyy", label: "MMMM.yyyy", hint: "Primer: februar.2025" },
    { value: "date:dd.mmmm.yyyy", label: "dd.MMMM.yyyy", hint: "Primer: 07.februar.2025" },
  ],
  number: [
    { value: "number:auto", label: "Automatski", hint: "" },
    { value: "number:int", label: "Ceo broj", hint: "Primer: 1.234" },
    { value: "number:2", label: "2 decimale", hint: "Primer: 1.234,56" },
    { value: "number:rsd", label: "RSD", hint: "Primer: 1.234,56 RSD" },
    { value: "number:eur", label: "€", hint: "Primer: 1.234,56 €" },
    { value: "number:usd", label: "$", hint: "Primer: 1.234,56 $" },
  ],
};

// ---------- DOM helpers ----------
function el(id) {
  return document.getElementById(id);
}

function setStatus(msg, kind = "info") {
  const s = el("status");
  if (!s) return;
  s.textContent = msg;
  s.className = `status ${kind}`;
  s.style.display = "block";
  clearTimeout(setStatus._t);
  setStatus._t = setTimeout(() => {
    s.style.display = "none";
  }, 3500);
}

// ---------- row helpers ----------
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
    // Očisti input - dozvoli tačke, zareze i cifre
    const cleanValue = String(v).replace(/[^\d.,-]/g, "");
    // Pretvori u broj - zameni zareze sa tačkama
    const n = Number(cleanValue.replace(/\./g, "").replace(",", "."));
    if (Number.isNaN(n)) return v;
    
    // Formatiranje broja: 1.234,56 format (tačka za hiljade, zarez za decimale)
    const formatNumber = (num, decimals = 0) => {
      const fixed = num.toFixed(decimals);
      const parts = fixed.split(".");
      // Dodaj tačku kao separator hiljada
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
      // Zameni decimalnu tačku sa zarezom
      return decimals > 0 ? parts[0] + "," + parts[1] : parts[0];
    };
    
    if (format === "number:int") return formatNumber(Math.round(n), 0);
    if (format === "number:2") return formatNumber(n, 2);
    if (format === "number:rsd") return formatNumber(n, 2) + " RSD";
    if (format === "number:eur") return formatNumber(n, 2) + " €";
    if (format === "number:usd") return formatNumber(n, 2) + " $";
    // Fallback za stari format
    if (format === "number:currency") return formatNumber(n, 2) + " RSD";
    
    return String(n);
  }

  if (type === "date") {
    // Meseci na srpskom (lowercase)
    const months = [
      "januar", "februar", "mart", "april", "maj", "jun",
      "jul", "avgust", "septembar", "oktobar", "novembar", "decembar"
    ];
    
    if (format === "date:today") {
      const d = new Date();
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const yyyy = d.getFullYear();
      return `${dd}.${mm}.${yyyy}`;
    }
    
    if (format === "date:mmmm.yyyy" || format === "date:dd.mmmm.yyyy") {
      // Parsiranje datuma (očekujemo dd.mm.yyyy ili yyyy-mm-dd)
      let d;
      if (v.includes(".")) {
        // dd.mm.yyyy format
        const parts = v.split(".");
        if (parts.length === 3) {
          d = new Date(parts[2], parts[1] - 1, parts[0]);
        }
      } else if (v.includes("-")) {
        // yyyy-mm-dd format
        d = new Date(v);
      } else {
        return v; // ne možemo parsirati
      }
      
      if (!d || isNaN(d.getTime())) return v;
      
      const monthName = months[d.getMonth()];
      const yyyy = d.getFullYear();
      
      if (format === "date:mmmm.yyyy") {
        return `${monthName}.${yyyy}`;
      } else {
        const dd = String(d.getDate()).padStart(2, "0");
        return `${dd}.${monthName}.${yyyy}`;
      }
    }
    
    return v;
  }

  // text formatting
  if (type === "text") {
    if (format === "text:upper") return v.toUpperCase();
    if (format === "text:lower") return v.toLowerCase();
    if (format === "text:title") {
      return v.replace(/\b\w/g, (l) => l.toUpperCase());
    }
  }

  return v;
}

function buildValueMap() {
  const map = new Map();
  for (const r of rows) {
    const key = normalizeKey(r.field);
    if (!key) continue;
    const raw = (r.value ?? "").trim();
    const formatted = applyFormat(r.type, r.format, raw);
    map.set(key, { raw, formatted });
  }
  return map;
}

// ---------- render rows ----------
function renderRows() {
  const container = el("rows");
  if (!container) return;

  container.innerHTML = "";

  if (rows.length === 0) {
    container.innerHTML = '<div class="empty-state">Nema polja. Klikni "+ Dodaj red".</div>';
    return;
  }

  rows.forEach((r, idx) => {
    const row = document.createElement("div");
    row.className = "row";  // PROMENJENO sa "table-row"
    if (idx === selectedRowIndex) row.classList.add("selected");

    // Field column
    const fieldCell = document.createElement("div");
    fieldCell.className = "cell";  // PROMENJENO sa "table-cell"
    const fieldInput = document.createElement("input");
    fieldInput.type = "text";
    fieldInput.placeholder = "Naziv polja";
    fieldInput.value = r.field || "";
    fieldInput.addEventListener("input", (e) => {
      r.field = e.target.value;
      saveStateToDocument();
    });
    fieldCell.appendChild(fieldInput);

    // Value column
    const valueCell = document.createElement("div");
    valueCell.className = "cell";  // PROMENJENO sa "table-cell"
    const valueInput = document.createElement("input");
    valueInput.type = "text";
    valueInput.placeholder = "Vrednost";
    valueInput.value = r.value || "";
    valueInput.addEventListener("input", (e) => {
      r.value = e.target.value;
      saveStateToDocument();
    });
    valueCell.appendChild(valueInput);

    // Actions column
    const actionsCell = document.createElement("div");
    actionsCell.className = "del";  // PROMENJENO sa "table-cell actions-cell"
    
    const btnEdit = document.createElement("button");
    btnEdit.innerHTML = "⚙";  // Jednostavan ASCII umesto emoji
    btnEdit.title = "Podešavanja (tip, format)";
    btnEdit.style.marginRight = "4px";
    btnEdit.style.width = "36px";
    btnEdit.style.height = "36px";
    btnEdit.style.border = "none";
    btnEdit.style.background = "#e0f2fe";
    btnEdit.style.color = "#0369a1";
    btnEdit.style.fontSize = "18px";
    btnEdit.style.cursor = "pointer";
    btnEdit.style.borderRadius = "6px";
    btnEdit.style.transition = "all 0.2s";
    btnEdit.addEventListener("mouseover", () => {
      btnEdit.style.background = "#bae6fd";
      btnEdit.style.transform = "scale(1.08)";
    });
    btnEdit.addEventListener("mouseout", () => {
      btnEdit.style.background = "#e0f2fe";
      btnEdit.style.transform = "scale(1)";
    });
    btnEdit.addEventListener("click", () => {
      selectedRowIndex = idx;
      openModal(r);
    });

    const btnDelete = document.createElement("button");
    btnDelete.innerHTML = "×";  // X umesto emoji
    btnDelete.title = "Obriši red";
    btnDelete.addEventListener("click", () => {
      if (confirm(`Obrisati polje "${r.field}"?`)) {
        rows.splice(idx, 1);
        if (selectedRowIndex === idx) selectedRowIndex = null;
        renderRows();
        saveStateToDocument();
      }
    });

    actionsCell.appendChild(btnEdit);
    actionsCell.appendChild(btnDelete);

    row.appendChild(fieldCell);
    row.appendChild(valueCell);
    row.appendChild(actionsCell);

    container.appendChild(row);
  });
}

// ---------- modal ----------
function openModal(row) {
  const modal = el("modal");
  const backdrop = el("modalBackdrop");
  const fieldName = el("modalFieldName");
  const formatSelect = el("formatSelect");
  const formatHint = el("formatHint");

  if (!modal || !backdrop) return;

  // Show modal
  modal.classList.remove("hidden");
  backdrop.classList.remove("hidden");

  // Set field name
  fieldName.textContent = row.field || "Neimenovano polje";

  // Set radio buttons
  const radios = document.querySelectorAll('input[name="ftype"]');
  radios.forEach((r) => {
    r.checked = r.value === (row.type || "text");
  });

  // Populate format dropdown
  updateFormatOptions(row.type || "text");
  formatSelect.value = row.format || "text:auto";
  updateFormatHint();

  // Event listeners
  radios.forEach((r) => {
    r.addEventListener("change", () => {
      updateFormatOptions(r.value);
      updateFormatHint();
    });
  });

  formatSelect.addEventListener("change", updateFormatHint);

  function updateFormatOptions(type) {
    formatSelect.innerHTML = "";
    const options = FORMAT_OPTIONS[type] || FORMAT_OPTIONS.text;
    options.forEach((opt) => {
      const option = document.createElement("option");
      option.value = opt.value;
      option.textContent = opt.label;
      formatSelect.appendChild(option);
    });
    formatSelect.value = options[0].value;
  }

  function updateFormatHint() {
    const type = document.querySelector('input[name="ftype"]:checked').value;
    const format = formatSelect.value;
    const options = FORMAT_OPTIONS[type] || FORMAT_OPTIONS.text;
    const opt = options.find((o) => o.value === format);
    formatHint.textContent = opt?.hint || "";
  }
}

function closeModal() {
  const modal = el("modal");
  const backdrop = el("modalBackdrop");
  if (modal) modal.classList.add("hidden");
  if (backdrop) backdrop.classList.add("hidden");
}

function saveModalChanges() {
  if (selectedRowIndex === null) return;

  const row = rows[selectedRowIndex];
  const type = document.querySelector('input[name="ftype"]:checked').value;
  const format = el("formatSelect").value;

  row.type = type;
  row.format = format;

  closeModal();
  renderRows();
  saveStateToDocument();
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
      rows = items;
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
  if (selectedRowIndex === null) {
    setStatus("Izaberi red u tabeli prvo.", "warn");
    return;
  }

  const r = rows[selectedRowIndex];
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
    cc.appearance = "BoundingBox";

    try {
      cc.cannotDelete = true;
      cc.cannotEdit = false;
    } catch (e) {
      // ignore
    }

    cc.insertText(token(key), Word.InsertLocation.replace);
    await context.sync();
    setStatus(`Dodato polje: ${key}`, "info");
  });

  await saveStateToDocument();
}

async function fillFieldsFromTable() {
  console.log("🔵 fillFieldsFromTable() POZVANA - NOVA VERZIJA (cc.insertText direktno)");

  const map = buildValueMap();

  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let filled = 0;

    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;

      const out = map.get(meta.key)?.formatted ?? token(meta.key);

      console.log(`  - Popunjavam polje "${meta.key}" sa: "${out}"`);
      cc.insertText(out, Word.InsertLocation.replace);

      filled++;
    }

    await context.sync();
    console.log(`✅ Popunjeno ${filled} polja`);
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

      cc.insertText(token(meta.key), Word.InsertLocation.replace);
      cleared++;
    }

    await context.sync();
    setStatus(`Očišćeno ${cleared} polja (placeholder {KEY} vraćen).`, "info");
  });

  await saveStateToDocument();
}

async function deleteControlsAndXml() {
  const confirmed = confirm(
    "PAŽNJA: Ova akcija će trajno obrisati sva polja i plugin podatke iz dokumenta.\n\n" +
      "Nakon brisanja, dokument neće više raditi sa ovim pluginom.\n\n" +
      "Da li želiš da nastaviš?"
  );

  if (!confirmed) {
    setStatus("Brisanje otkazano.", "info");
    return;
  }

  const map = buildValueMap();

  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    let removed = 0;
    for (const cc of ccs.items) {
      const meta = parseTag(cc.tag || "");
      if (!meta) continue;

      const out = map.get(meta.key)?.formatted ?? "";

      cc.insertText(out, Word.InsertLocation.replace);
      try {
        cc.delete(false);
      } catch (e) {
        // ignore
      }
      removed++;
    }
    await context.sync();
  });

  await deleteSavedStateFromDocument();

  rows = [];
  selectedRowIndex = null;
  renderRows();

  setStatus("Dokument očišćen: polja i plugin podaci su uklonjeni.", "info");
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

  setStatus(`Eksportovano ${lines.length} polja u CSV.`, "info");
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
        newRows.push({ field, value, type: "text", format: "text:auto" });
      }
    }

    if (newRows.length) {
      rows = newRows;
      selectedRowIndex = null;
      renderRows();
      await saveStateToDocument();
      setStatus(`Importovano ${newRows.length} polja iz CSV.`, "info");
    }
  };

  input.click();
}

// ---------- wiring ----------
function bindUi() {
  const btnInsert = el("btnInsert");
  const btnFill = el("btnFill");
  const btnClear = el("btnClear");
  const btnDelete = el("btnDelete");
  const btnAddRow = el("btnAddRow");
  const btnExportCSV = el("btnExportCSV");
  const btnImportCSV = el("btnImportCSV");

  const btnModalClose = el("btnModalClose");
  const btnModalCancel = el("btnModalCancel");
  const btnModalOk = el("btnModalOk");
  const modalBackdrop = el("modalBackdrop");

  if (btnInsert) btnInsert.addEventListener("click", insertFieldAtSelection);
  if (btnFill) btnFill.addEventListener("click", fillFieldsFromTable);
  if (btnClear) btnClear.addEventListener("click", clearFieldsKeepControls);
  if (btnDelete) btnDelete.addEventListener("click", deleteControlsAndXml);
  if (btnExportCSV) btnExportCSV.addEventListener("click", exportCSV);
  if (btnImportCSV) btnImportCSV.addEventListener("click", importCSV);

  if (btnAddRow) {
    btnAddRow.addEventListener("click", () => {
      rows.push({ field: "", value: "", type: "text", format: "text:auto" });
      renderRows();
      saveStateToDocument();
    });
  }

  if (btnModalClose) btnModalClose.addEventListener("click", closeModal);
  if (btnModalCancel) btnModalCancel.addEventListener("click", closeModal);
  if (btnModalOk) btnModalOk.addEventListener("click", saveModalChanges);
  if (modalBackdrop) modalBackdrop.addEventListener("click", closeModal);
}

Office.onReady(async () => {
  try {
    await loadStateFromDocument();
  } catch (e) {
    console.error("Load state error:", e);
  }

  renderRows();
  bindUi();
});