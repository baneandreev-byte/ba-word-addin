/* global Office, Word */

// ============================================
// VERZIJA: 2026-02-26 - V64
// ============================================
console.log("🔧 BA Word Add-in VERZIJA: 2026-02-26 - V63");
console.log("✅ NOVO: Tabele koriste hidden tag pattern umesto Content Controls");
console.log("✅ Placeholder: obična tabela sa [BA:04]/[BA:05]/[BA:061]/[BA:062]/[BA:08] u prvoj ćeliji");

let rows = [];
let selectedRowIndex = null;

// Drag & Drop state (koristi ID umesto index za stabilnost)
let draggedElement = null;
let draggedId = null;

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

// ---------- Drag & Drop handlers ----------
function handleDragStart(e) {
  const handle = e.currentTarget; // handle element
  const rowEl = handle.closest(".row");
  
  draggedElement = rowEl;
  draggedId = handle.dataset.id;
  
  rowEl.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', draggedId); // stabilnije od text/html
}

function handleDragOver(e) {
  if (e.preventDefault) {
    e.preventDefault();
  }
  
  e.dataTransfer.dropEffect = 'move';
  
  // Visual feedback - show drop indicator
  const targetRow = e.target.closest('.row');
  if (targetRow && targetRow !== draggedElement) {
    // Remove drag-over from all rows first
    document.querySelectorAll('.row').forEach(r => {
      if (r !== targetRow) r.classList.remove('drag-over');
    });
    targetRow.classList.add('drag-over');
  }
  
  return false;
}

function handleDragLeave(e) {
  const targetRow = e.target.closest('.row');
  if (targetRow) {
    targetRow.classList.remove('drag-over');
  }
}

function handleDrop(e) {
  if (e.stopPropagation) {
    e.stopPropagation();
  }
  
  const targetRow = e.target.closest('.row');
  if (!targetRow || targetRow === draggedElement) {
    return false;
  }
  
  const targetId = targetRow.dataset.id;
  
  // Pronađi indekse u rows array-u pomoću ID-a (stabilno)
  const fromIndex = rows.findIndex(r => r.id === draggedId);
  const toIndex = rows.findIndex(r => r.id === targetId);
  
  if (fromIndex === -1 || toIndex === -1) return false;
  
  // Reorder rows array
  const [movedItem] = rows.splice(fromIndex, 1);
  rows.splice(toIndex, 0, movedItem);
  
  // Update selected index if needed - guard protiv null/undefined
  if (selectedRowIndex !== null && selectedRowIndex !== undefined) {
    if (selectedRowIndex === fromIndex) {
      selectedRowIndex = toIndex;
    } else if (fromIndex < selectedRowIndex && toIndex >= selectedRowIndex) {
      selectedRowIndex--;
    } else if (fromIndex > selectedRowIndex && toIndex <= selectedRowIndex) {
      selectedRowIndex++;
    }
  }
  
  // Re-render and save
  renderRows();
  saveStateToDocument();
  
  // Show status
  setStatus(`Polje "${movedItem.field}" premešteno.`, "info");
  
  return false;
}

function handleDragEnd() {
  if (draggedElement) {
    draggedElement.classList.remove('dragging');
  }
  
  // Remove all drag-over classes
  document.querySelectorAll('.row').forEach(row => {
    row.classList.remove('drag-over');
  });
  
  draggedElement = null;
  draggedId = null;
}

// ---------- render rows ----------
function renderRows() {
  const container = el("rows");
  if (!container) return;

  // Sačuvaj trenutni focus
  const activeElement = document.activeElement;
  const wasFieldInput = activeElement && activeElement.placeholder === "Naziv polja";
  const wasValueInput = activeElement && activeElement.placeholder === "Vrednost";
  const focusedRowIndex = wasFieldInput || wasValueInput ? 
    Array.from(container.querySelectorAll('.row')).indexOf(activeElement.closest('.row')) : -1;

  container.innerHTML = "";

  if (rows.length === 0) {
    container.innerHTML = '<div class="empty-state">Nema polja. Klikni "+ Dodaj red".</div>';
    return;
  }

  rows.forEach((r, idx) => {
    // Osiguraj da svaki red ima ID
    if (!r.id) r.id = crypto.randomUUID();
    
    const row = document.createElement("div");
    row.className = "row";
    if (idx === selectedRowIndex) row.classList.add("selected");
    
    // Red NIJE draggable - samo handle jeste
    row.draggable = false;
    row.dataset.id = r.id;
    row.dataset.index = idx; // Zadrži index za backward compatibility
    
    // Drag event listeners na RED (drop target)
    row.addEventListener('dragover', handleDragOver);
    row.addEventListener('dragleave', handleDragLeave);
    row.addEventListener('drop', handleDrop);

    // Click handler na ceo red - selektuje red za ubacivanje
    row.addEventListener("click", (e) => {
      // Don't select if clicking drag handle
      if (e.target.closest('.drag-handle')) return;
      selectedRowIndex = idx;
      renderRows();
    });

    // Drag handle - SAMO handle je draggable
    const dragHandle = document.createElement("div");
    dragHandle.className = "drag-handle";
    dragHandle.innerHTML = "⋮⋮";
    dragHandle.title = "Prevuci za premeštanje";
    dragHandle.draggable = true;
    dragHandle.dataset.id = r.id;
    
    // Drag event listeners na HANDLE (drag source)
    dragHandle.addEventListener('dragstart', handleDragStart);
    dragHandle.addEventListener('dragend', handleDragEnd);

    // Field column
    const fieldCell = document.createElement("div");
    fieldCell.className = "cell";
    const fieldInput = document.createElement("input");
    fieldInput.type = "text";
    fieldInput.placeholder = "Naziv polja";
    fieldInput.value = r.field || "";
    fieldInput.addEventListener("input", (e) => {
      r.field = e.target.value;
      saveStateToDocument();
    });
    // Selektuj red kada se klikne na input
    fieldInput.addEventListener("click", (e) => {
      e.stopPropagation(); // Spreči dupli event
      if (selectedRowIndex !== idx) {
        selectedRowIndex = idx;
        renderRows();
      }
    });
    // Selektuj red kada input dobije focus
    fieldInput.addEventListener("focus", () => {
      if (selectedRowIndex !== idx) {
        selectedRowIndex = idx;
        renderRows();
      }
    });
    fieldCell.appendChild(fieldInput);

    // Value column
    const valueCell = document.createElement("div");
    valueCell.className = "cell";
    const valueInput = document.createElement("input");
    valueInput.type = "text";
    valueInput.placeholder = "Vrednost";
    valueInput.value = r.value || "";
    valueInput.addEventListener("input", (e) => {
      r.value = e.target.value;
      saveStateToDocument();
    });
    // Selektuj red kada se klikne na input
    valueInput.addEventListener("click", (e) => {
      e.stopPropagation(); // Spreči dupli event
      if (selectedRowIndex !== idx) {
        selectedRowIndex = idx;
        renderRows();
      }
    });
    // Selektuj red kada input dobije focus
    valueInput.addEventListener("focus", () => {
      if (selectedRowIndex !== idx) {
        selectedRowIndex = idx;
        renderRows();
      }
    });
    valueCell.appendChild(valueInput);

    // Actions column
    const actionsCell = document.createElement("div");
    actionsCell.className = "del";
    
    const btnEdit = document.createElement("button");
    btnEdit.innerHTML = "⚙";
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
    btnEdit.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedRowIndex = idx;
      openModal(r);
    });

    const btnDelete = document.createElement("button");
    btnDelete.innerHTML = "×";
    btnDelete.title = "Obriši red";
    btnDelete.addEventListener("click", (e) => {
      e.stopPropagation();
      if (confirm(`Obrisati polje "${r.field}"?`)) {
        rows.splice(idx, 1);
        if (selectedRowIndex === idx) selectedRowIndex = null;
        renderRows();
        saveStateToDocument();
      }
    });

    actionsCell.appendChild(btnEdit);
    actionsCell.appendChild(btnDelete);

    row.appendChild(dragHandle);
    row.appendChild(fieldCell);
    row.appendChild(valueCell);
    row.appendChild(actionsCell);

    container.appendChild(row);
  });

  // Vrati focus ako je bio aktivan
  if (focusedRowIndex >= 0 && focusedRowIndex < rows.length) {
    const allRows = container.querySelectorAll('.row');
    const targetRow = allRows[focusedRowIndex];
    if (targetRow) {
      if (wasFieldInput) {
        const fieldInput = targetRow.querySelector('input[placeholder="Naziv polja"]');
        if (fieldInput) setTimeout(() => fieldInput.focus(), 0);
      } else if (wasValueInput) {
        const valueInput = targetRow.querySelector('input[placeholder="Vrednost"]');
        if (valueInput) setTimeout(() => valueInput.focus(), 0);
      }
    }
  }
}

// ---------- modal ----------
function openModal(row) {
  const modal = el("modal");
  const backdrop = el("modalBackdrop");
  const fieldNameSpan = el("modalFieldName");
  const formatSelect = el("formatSelect");
  const formatHint = el("formatHint");

  if (!modal || !backdrop) return;

  // Display field name
  if (fieldNameSpan) {
    fieldNameSpan.textContent = row.field || "(bez naziva)";
  }

  // Set radio button for type
  const radios = document.querySelectorAll('input[name="ftype"]');
  radios.forEach((r) => {
    r.checked = r.value === row.type;
  });

  // Populate format dropdown
  updateFormatOptions(row.type, row.format);

  // Listen to type changes
  radios.forEach((r) => {
    r.addEventListener("change", () => {
      updateFormatOptions(r.value, null);
    });
  });

  // Listen to format changes for hint
  if (formatSelect) {
    formatSelect.addEventListener("change", () => {
      const selectedOption = formatSelect.options[formatSelect.selectedIndex];
      if (formatHint && selectedOption) {
        formatHint.textContent = selectedOption.getAttribute("data-hint") || "";
      }
    });
  }

  // Show modal
  modal.classList.remove("hidden");
  backdrop.classList.remove("hidden");
}

function closeModal() {
  const modal = el("modal");
  const backdrop = el("modalBackdrop");
  if (modal) modal.classList.add("hidden");
  if (backdrop) backdrop.classList.add("hidden");
}

// ---------- Delete Confirm Modal ----------
function showDeleteConfirmModal() {
  const modal = el("deleteModal");
  const backdrop = el("deleteModalBackdrop");
  if (!modal || !backdrop) {
    console.error("❌ Delete modal elementi ne postoje u HTML-u!");
    return;
  }
  
  modal.classList.remove("hidden");
  backdrop.classList.remove("hidden");
}

function closeDeleteModal() {
  const modal = el("deleteModal");
  const backdrop = el("deleteModalBackdrop");
  if (modal) modal.classList.add("hidden");
  if (backdrop) backdrop.classList.add("hidden");
}

function updateFormatOptions(type, currentFormat) {
  const formatSelect = el("formatSelect");
  const formatHint = el("formatHint");
  if (!formatSelect) return;

  formatSelect.innerHTML = "";

  const opts = FORMAT_OPTIONS[type] || FORMAT_OPTIONS.text;
  opts.forEach((opt) => {
    const option = document.createElement("option");
    option.value = opt.value;
    option.textContent = opt.label;
    option.setAttribute("data-hint", opt.hint);
    if (currentFormat && opt.value === currentFormat) {
      option.selected = true;
    }
    formatSelect.appendChild(option);
  });

  // Set hint for selected option
  if (formatHint) {
    const selectedOption = formatSelect.options[formatSelect.selectedIndex];
    formatHint.textContent = selectedOption ? selectedOption.getAttribute("data-hint") || "" : "";
  }
}

function saveModalChanges() {
  if (selectedRowIndex === null) return;

  const row = rows[selectedRowIndex];
  
  // Get selected type
  const checkedRadio = document.querySelector('input[name="ftype"]:checked');
  if (checkedRadio) {
    row.type = checkedRadio.value;
  }

  // Get selected format
  const formatSelect = el("formatSelect");
  if (formatSelect) {
    row.format = formatSelect.value;
  }

  closeModal();
  renderRows();
  saveStateToDocument();
  setStatus(`Ažurirano: ${row.field}`, "info");
}

// ---------- XML state ----------
const XML_NS = "http://biroa.rs/word-addin/state";

function xmlEscape(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function buildStateXml() {
  let xml = `<state xmlns="${XML_NS}">`;
  for (const r of rows) {
    const f = xmlEscape(r.field || "");
    const v = xmlEscape(r.value || "");
    const t = xmlEscape(r.type || "text");
    const fmt = xmlEscape(r.format || "text:auto");
    xml += `<item field="${f}" value="${v}" type="${t}" format="${fmt}"/>`;
  }
  xml += "</state>";
  return xml;
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
  console.log("🔵 fillFieldsFromTable() POZVANA");

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

// ============================================
// FIX: deleteControlsAndXml - Custom Modal umesto confirm()
// ============================================
/**
 * FAZA 1: Mapiranje svih BA_FIELD kontrola
 * Analizira dokument i vraća listu kontrola koje će biti obrisane
 */
async function mapControlsForDeletion() {
  console.log("🔄 FAZA 1: Mapiranje kontrola za brisanje...");
  console.log("=".repeat(60));
  
  const controlsToDelete = [];
  
  await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    const totalControls = contentControls.items.length;
    console.log(`📊 Ukupno content controls u dokumentu: ${totalControls}`);

    if (totalControls === 0) {
      console.log("ℹ️ Nema content control-a u dokumentu");
      return;
    }

    // Učitaj properties za sve kontrole
    for (const cc of contentControls.items) {
      cc.load("tag,text,title");
    }
    await context.sync();
    console.log("✅ Properties učitane za sve kontrole");

    // Analiziraj svaku kontrolu
    console.log("\n📋 Detaljno mapiranje:");
    console.log("-".repeat(60));
    
    for (let i = 0; i < contentControls.items.length; i++) {
      const cc = contentControls.items[i];
      const tag = cc.tag || "";
      const title = cc.title || "(bez naslova)";
      const text = cc.text || "";
      
      console.log(`\n[${i}] Kontrola:`);
      console.log(`    Title: "${title}"`);
      console.log(`    Tag: "${tag}"`);
      console.log(`    Text: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"`);
      
      const meta = parseTag(tag);
      
      if (!meta) {
        console.log(`    ⏭️ PRESKAČEM - nije BA_FIELD kontrola`);
        continue;
      }
      
      console.log(`    ✅ BIĆE OBRISAN - BA_FIELD: ${meta.key}`);
      
      // Dodaj u listu za brisanje
      controlsToDelete.push({
        index: i,
        key: meta.key,
        type: meta.type,
        format: meta.format,
        text: text,
        title: title
      });
    }

    console.log("-".repeat(60));
    console.log(`\n📊 Rezime mapiranja:`);
    console.log(`   Total kontrola: ${totalControls}`);
    console.log(`   BA_FIELD kontrola: ${controlsToDelete.length}`);
    console.log(`   Preskočeno: ${totalControls - controlsToDelete.length}`);
    console.log("=".repeat(60));
  });

  return controlsToDelete;
}

/**
 * FAZA 2: Prikaz confirmation dialog-a sa listom kontrola
 */
function showDeleteConfirmationWithList(controlsList) {
  console.log("\n💬 Prikazujem confirmation dialog sa listom kontrola...");
  
  // Pripremi HTML listu kontrola
  let listHtml = '';
  if (controlsList.length === 0) {
    listHtml = '<p style="text-align: center; color: #9ca3af; font-style: italic;">Nema BA_FIELD kontrola za brisanje</p>';
  } else {
    listHtml = '<div style="max-height: 300px; overflow-y: auto; margin: 16px 0;">';
    controlsList.forEach((ctrl, idx) => {
      const truncatedText = ctrl.text.length > 60 
        ? ctrl.text.substring(0, 60) + '...' 
        : ctrl.text;
      
      listHtml += `
        <div style="
          padding: 12px;
          margin-bottom: 8px;
          border: 1px solid #e5e7eb;
          border-radius: 6px;
          background: #f9fafb;
        ">
          <div style="font-weight: 600; color: #1f2937; margin-bottom: 4px;">
            ${idx + 1}. ${ctrl.key}
          </div>
          <div style="font-size: 12px; color: #6b7280; font-style: italic;">
            → "${truncatedText || '(prazno)'}"
          </div>
        </div>
      `;
    });
    listHtml += '</div>';
  }

  // Ažuriraj modal body sa listom
  const modal = el("deleteModal");
  if (modal) {
    const modalBody = modal.querySelector(".modal-body");
    if (modalBody) {
      modalBody.innerHTML = `
        <p style="margin-bottom: 16px; color: #6b7280; line-height: 1.6;">
          Pronađeno je <strong style="color: #1f2937;">${controlsList.length}</strong> aktivnih polja koja će biti uklonjena:
        </p>
        ${listHtml}
        <p style="margin-top: 16px; margin-bottom: 0; color: #1f2937; font-weight: 600; text-align: center; padding: 12px; background: #fef3c7; border-radius: 6px; border: 1px solid #fbbf24;">
          ⚠️ Tekst iz svakog polja će biti sačuvan u dokumentu
        </p>
      `;
    }
  }

  // Prikaži modal
  showDeleteConfirmModal();
}

async function deleteControlsAndXml() {
  try {
    console.log("🔴 deleteControlsAndXml() - POČETAK");
    console.log("=".repeat(60));
    
    // FAZA 1: Mapiraj sve kontrole
    const controlsList = await mapControlsForDeletion();
    
    if (controlsList.length === 0) {
      setStatus("Nema aktivnih polja za brisanje.", "info");
      return;
    }
    
    // Sačuvaj listu u globalnu promenljivu za performDelete
    window._controlsToDelete = controlsList;
    
    // FAZA 2: Prikaži confirmation dialog sa listom
    showDeleteConfirmationWithList(controlsList);
    
  } catch (error) {
    console.error("❌ Greška pri mapiranju kontrola:", error);
    console.error("❌ Stack:", error.stack);
    setStatus("Greška pri analizi polja.", "error");
  }
}

async function performDelete() {
  try {
    console.log("🔄 FAZA 3: Izvršavanje brisanja nakon potvrde (STABILNA strategija)...");
    console.log("=".repeat(60));

    const controlsList = window._controlsToDelete || [];
    const keysToDelete = new Set(
      controlsList
        .map((c) => String(c?.key || "").trim())
        .filter((k) => k.length > 0)
    );

    let removed = 0;

    await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      // Učitaj minimalno neophodne podatke
      contentControls.load("items/tag,text,cannotDelete");
      await context.sync();

      const all = contentControls.items || [];
      console.log(`📊 Trenutno content controls u dokumentu: ${all.length}`);

      // Filtriraj BA_FIELD kontrole (po tag-u), bez oslanjanja na indekse
      const targets = [];
      for (const cc of all) {
        const meta = parseTag(cc.tag || "");
        if (!meta) continue;
        // Ako imamo listu ključeva iz modala, briši samo te; u suprotnom briši sve BA_FIELD
        if (keysToDelete.size > 0 && !keysToDelete.has(meta.key)) continue;
        targets.push({ cc, meta });
      }

      if (targets.length === 0) {
        console.log("ℹ️ Nema BA_FIELD kontrola za brisanje (u trenutnom stanju dokumenta).");
        return;
      }

      console.log(`🧹 Biće obrisano ${targets.length} BA_FIELD kontrola (wrapper), tekst ostaje.`);

      // 1) Otključaj sve (ako je potrebno) u jednoj turi
      for (const t of targets) {
        if (t.cc.cannotDelete) t.cc.cannotDelete = false;
      }

      // 2) Obriši wrapper, zadrži sadržaj (keepContent=true)
      for (const t of targets) {
        t.cc.delete(true);
      }

      await context.sync();
      removed = targets.length;

      console.log(`✅ ZAVRŠENO: Obrisano ${removed} kontrola (wrapper), sadržaj sačuvan.`);
    });

    if (removed === 0) {
      setStatus("Nema polja za brisanje.", "info");
      closeDeleteModal();
      return;
    }

    // Obriši XML state
    try {
      await deleteSavedStateFromDocument();
      console.log("✅ XML state obrisan");
    } catch (error) {
      console.warn("⚠️ XML state greška (nije kritično):", error);
    }

    // Očisti lokalne podatke
    rows = [];
    selectedRowIndex = null;
    renderRows();

    // Očisti globalnu listu
    window._controlsToDelete = null;

    setStatus(`Dokument očišćen: ${removed} polja uklonjeno.`, "info");
    closeDeleteModal();
    console.log("=".repeat(60));
  } catch (error) {
    console.error("❌ GREŠKA pri brisanju:", error);
    console.error("❌ Stack:", error.stack);
    setStatus("Greška pri brisanju polja. Vidi konzolu.", "error");
    closeDeleteModal();
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
        newRows.push({ 
          id: crypto.randomUUID(), 
          field, 
          value, 
          type: "text", 
          format: "text:auto" 
        });
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

// ============================================
// TEMPLATE MANAGER (V56 - GitHub + Struke Picker)
// ============================================

// GitHub konfiguracija
const GITHUB_CONFIG = {
  baseUrl: "https://raw.githubusercontent.com/baneandreev-byte/BiroA-templates-test/main",
  branches: [
    {
      id: "01 IDR", label: "IDR – Idejno rešenje",
      files: [
        { name: "00 IDR Glavna sveska.dotx", isGlavna: true },
        { name: "01 IDR Sveska projekta.dotx", isGlavna: false }
      ]
    },
    {
      id: "02 PGD", label: "PGD – Projekat za građevinsku dozvolu",
      files: [
        { name: "00 PGD Glavna sveska.dotx", isGlavna: true },
        { name: "01 PGD Sveska projekta.dotx", isGlavna: false }
      ]
    },
    {
      id: "03 PZI", label: "PZI – Projekat za izvođenje",
      files: [
        { name: "00 PZI Glavna sveska.dotx", isGlavna: true },
        { name: "01 PZI Sveska projetka.dotx", isGlavna: false }
      ]
    },
    {
      id: "04 TK", label: "TK – Tehnička kontrola",
      files: [
        { name: "00 TK Glavna sveska.dotx", isGlavna: true },
        { name: "01 TK Sveska projekta.docx", isGlavna: false }
      ]
    }
  ]
};

// Fiksna lista struka iz Excel-a
// Grupe: oznaka grupe je prefiks pre "/" ili cela oznaka ako nema "/"
const STRUKE_LIST = [
  { oznaka: "00", naziv: "Glavna sveska", grupa: "00" },
  { oznaka: "01", naziv: "Arhitektura", grupa: "01" },
  { oznaka: "02/1", naziv: "Konstrukcija", grupa: "02" },
  { oznaka: "02/2", naziv: "Saobraćajna konstrukcija", grupa: "02" },
  { oznaka: "03/1", naziv: "Hidrotehničke instalacije", grupa: "03" },
  { oznaka: "04/1", naziv: "Elektroenergetske instalacije", grupa: "04" },
  { oznaka: "04/2", naziv: "Trafo stanica", grupa: "04" },
  { oznaka: "05/1", naziv: "Signalne instalacije", grupa: "05" },
  { oznaka: "05/2", naziv: "Stabilni sistem za dojavu požara", grupa: "05" },
  { oznaka: "05/3", naziv: "Video nadzor", grupa: "05" },
  { oznaka: "06/1", naziv: "Termotehničke instalacije", grupa: "06" },
  { oznaka: "06/2", naziv: "Putnički lift", grupa: "06" },
  { oznaka: "06/3", naziv: "Autolift", grupa: "06" },
  { oznaka: "06/4", naziv: "Stabilni sistem za gašenje požara", grupa: "06" },
  { oznaka: "06/5", naziv: "Ventilacija i nadpritisak garaže", grupa: "06" },
  { oznaka: "07", naziv: "Tehnologija", grupa: "07" },
  { oznaka: "08", naziv: "Saobraćajna signalizacija", grupa: "08" },
  { oznaka: "09", naziv: "Spoljno uređenje", grupa: "09" },
  { oznaka: "10/1", naziv: "Pripremni radovi - Projekat rušenja", grupa: "10" },
  { oznaka: "10/2", naziv: "Pripremni radovi - Projekat obezbeđenja temeljne jame", grupa: "10" },
  { oznaka: "GEO", naziv: "Elaborat geomehanike", grupa: "GEO" },
  { oznaka: "EE", naziv: "Elaborat energetske efikasnosti", grupa: "EE" },
];

// State za picker
let _pickerSelectedBranch = null;
let _pickerSelectedFile = null;
// Struke state: { grupa -> [ { naziv, checked, custom } ] }
let _strukeState = {};

let templates = [];
let editingTemplateId = null;

// ---------- GitHub Helpers ----------

function buildGitHubRawUrl(branchId, fileName) {
  return `${GITHUB_CONFIG.baseUrl}/${branchId.split("/").map(encodeURIComponent).join("/")}/${encodeURIComponent(fileName)}`;
}

async function downloadFileContent(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`GitHub download greška: ${response.status}`);
  return await response.arrayBuffer();
}

// Placeholder - ne treba API poziv pri startu
async function loadTemplatesFromSharePoint() {
  templates = [];
}

// ---------- Struke logika ----------

// Inicijalizuj state struka (sve unchecked, bez custom)
function initStrukeState() {
  _strukeState = {};
  for (const s of STRUKE_LIST) {
    if (s.grupa === "00") continue; // Glavna sveska se ne bira
    if (!_strukeState[s.grupa]) {
      _strukeState[s.grupa] = [];
    }
    _strukeState[s.grupa].push({ naziv: s.naziv, checked: false, custom: false });
  }
}

// Izračunaj konačne oznake za sve čekirane stavke
function computeOznake() {
  const result = [];

  // Uvek ide Glavna sveska kao 00
  result.push({ oznaka: "00", naziv: "Glavna sveska" });

  for (const [grupa, stavke] of Object.entries(_strukeState)) {
    const checked = stavke.filter(s => s.checked);
    if (checked.length === 0) continue;

    if (checked.length === 1) {
      // Jedna stavka u grupi → bez podbroja
      result.push({ oznaka: grupa, naziv: checked[0].naziv });
    } else {
      // Više stavki → grupa/1, grupa/2...
      checked.forEach((s, i) => {
        result.push({ oznaka: `${grupa}/${i + 1}`, naziv: s.naziv });
      });
    }
  }

  return result;
}

// ============================================
// GITHUB TEMPLATE PICKER MODAL
// ============================================

function openGitHubTemplateModal() {
  let backdrop = el("githubTemplateBackdrop");
  if (!backdrop) {
    backdrop = document.createElement("div");
    backdrop.id = "githubTemplateBackdrop";
    backdrop.style.cssText = `
      position:fixed; inset:0; background:rgba(0,0,0,0.5);
      z-index:1000; display:flex; align-items:center; justify-content:center;
    `;
    backdrop.addEventListener("click", e => {
      if (e.target === backdrop) closeGitHubTemplateModal();
    });

    const modal = document.createElement("div");
    modal.id = "githubTemplateModal";
    modal.style.cssText = `
      background:#fff; border-radius:10px; width:420px; max-width:95vw;
      max-height:90vh; display:flex; flex-direction:column;
      box-shadow:0 8px 32px rgba(0,0,0,0.2);
      font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
    `;

    const header = document.createElement("div");
    header.id = "githubTemplateHeader";
    header.style.cssText = `
      background:#1d4ed8; color:#fff; padding:14px 18px;
      display:flex; align-items:center; justify-content:space-between;
      border-radius:10px 10px 0 0; flex-shrink:0;
    `;

    const body = document.createElement("div");
    body.id = "githubTemplateBody";
    body.style.cssText = "padding:18px; overflow-y:auto; flex:1;";

    const footer = document.createElement("div");
    footer.id = "githubTemplateFooter";
    footer.style.cssText = `
      padding:12px 18px; border-top:1px solid #e5e7eb;
      display:flex; gap:8px; justify-content:flex-end; flex-shrink:0;
    `;

    modal.appendChild(header);
    modal.appendChild(body);
    modal.appendChild(footer);
    backdrop.appendChild(modal);
    document.body.appendChild(backdrop);
  }

  _pickerSelectedBranch = null;
  _pickerSelectedFile = null;
  initStrukeState();
  renderPickerStep1();
  el("githubTemplateBackdrop").style.display = "flex";
}

function closeGitHubTemplateModal() {
  const backdrop = el("githubTemplateBackdrop");
  if (backdrop) backdrop.style.display = "none";
}

function setPickerHeader(title, showBack, onBack) {
  const header = el("githubTemplateHeader");
  if (!header) return;
  header.innerHTML = `
    <div style="display:flex; align-items:center; gap:10px;">
      ${showBack ? `<button id="pickerBackBtn" style="background:rgba(255,255,255,0.2);border:none;color:#fff;font-size:16px;cursor:pointer;border-radius:4px;padding:2px 8px;">←</button>` : ""}
      <span style="font-weight:700; font-size:15px;">${title}</span>
    </div>
    <button id="pickerCloseBtn" style="background:none;border:none;color:#fff;font-size:22px;cursor:pointer;line-height:1;">×</button>
  `;
  document.getElementById("pickerCloseBtn").addEventListener("click", closeGitHubTemplateModal);
  if (showBack) document.getElementById("pickerBackBtn").addEventListener("click", onBack);
}

function setPickerFooter(buttons) {
  const footer = el("githubTemplateFooter");
  if (!footer) return;
  footer.innerHTML = "";
  buttons.forEach(b => {
    const btn = document.createElement("button");
    btn.textContent = b.label;
    btn.style.cssText = b.primary
      ? "padding:8px 20px;background:#1d4ed8;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;"
      : "padding:8px 18px;border:1px solid #d1d5db;border-radius:6px;background:#fff;cursor:pointer;font-size:13px;color:#374151;";
    btn.addEventListener("click", b.onClick);
    footer.appendChild(btn);
  });
}

// KORAK 1: Izbor grane
function renderPickerStep1() {
  setPickerHeader("📁 Izaberi vrstu projekta", false, null);
  const body = el("githubTemplateBody");
  body.innerHTML = "";

  GITHUB_CONFIG.branches.forEach(branch => {
    const btn = document.createElement("button");
    btn.style.cssText = `
      display:block; width:100%; text-align:left; padding:12px 14px;
      margin-bottom:8px; border:1px solid #e5e7eb; border-radius:8px;
      background:#f9fafb; cursor:pointer; font-size:13px; color:#1f2937;
      transition:all 0.15s;
    `;
    btn.innerHTML = `
      <div style="font-weight:700;color:#1d4ed8;">${branch.id.replace(/^\d+\s/, "")}</div>
      <div style="font-size:12px;color:#6b7280;margin-top:2px;">${branch.label}</div>
    `;
    btn.addEventListener("mouseenter", () => { btn.style.background="#eff6ff"; btn.style.borderColor="#93c5fd"; });
    btn.addEventListener("mouseleave", () => { btn.style.background="#f9fafb"; btn.style.borderColor="#e5e7eb"; });
    btn.addEventListener("click", () => {
      _pickerSelectedBranch = branch;
      renderPickerStep2();
    });
    body.appendChild(btn);
  });

  setPickerFooter([
    { label: "Otkaži", onClick: closeGitHubTemplateModal }
  ]);
}

// KORAK 2: Izbor sveske
function renderPickerStep2() {
  const branch = _pickerSelectedBranch;
  setPickerHeader(`📂 ${branch.label}`, true, renderPickerStep1);
  const body = el("githubTemplateBody");
  body.innerHTML = `<p style="margin:0 0 14px;color:#374151;font-size:13px;font-weight:600;">Izaberi svesku:</p>`;

  branch.files.forEach(file => {
    const btn = document.createElement("button");
    btn.style.cssText = `
      display:block; width:100%; text-align:left; padding:12px 14px;
      margin-bottom:8px; border:1px solid #e5e7eb; border-radius:8px;
      background:#f9fafb; cursor:pointer; font-size:13px; color:#1f2937;
      transition:all 0.15s;
    `;
    btn.innerHTML = `<span style="margin-right:8px;">${file.isGlavna ? "📋" : "📄"}</span>${file.name}`;
    btn.addEventListener("mouseenter", () => { btn.style.background="#eff6ff"; btn.style.borderColor="#93c5fd"; });
    btn.addEventListener("mouseleave", () => { btn.style.background="#f9fafb"; btn.style.borderColor="#e5e7eb"; });
    btn.addEventListener("click", () => {
      _pickerSelectedFile = file;
      openTemplateFromGitHub();
    });
    body.appendChild(btn);
  });

  setPickerFooter([
    { label: "Otkaži", onClick: closeGitHubTemplateModal }
  ]);
}

// KORAK 3: Izbor struka (samo za Glavnu svesku)
function renderPickerStep3() {
  setPickerHeader("🏗️ Izaberi struke projekta", true, renderPickerStep2);
  const body = el("githubTemplateBody");
  body.innerHTML = "";

  // Grupiši struke po grupi
  const grupe = {};
  for (const s of STRUKE_LIST) {
    if (s.grupa === "00") continue;
    if (!grupe[s.grupa]) grupe[s.grupa] = { naziv: s.naziv.split(" ")[0], stavke: [] };
    grupe[s.grupa].stavke.push(s);
  }

  for (const [grupa, info] of Object.entries(grupe)) {
    const state = _strukeState[grupa] || [];

    const groupDiv = document.createElement("div");
    groupDiv.style.cssText = `
      border:1px solid #e5e7eb; border-radius:8px; margin-bottom:10px; overflow:hidden;
    `;

    // Zaglavlje grupe
    const groupHeader = document.createElement("div");
    groupHeader.style.cssText = `
      background:#f3f4f6; padding:8px 12px; display:flex; 
      align-items:center; gap:8px; font-size:13px; font-weight:600; color:#1f2937;
    `;
    groupHeader.innerHTML = `<span style="color:#6b7280;font-size:11px;min-width:28px;">${grupa}</span>`;

    const groupBody = document.createElement("div");
    groupBody.id = `strukaGroup_${grupa}`;
    groupBody.style.cssText = "padding:4px 0;";

    // Stavke u grupi
    state.forEach((stavka, idx) => {
      const row = createStrukaRow(grupa, idx, stavka);
      groupBody.appendChild(row);
    });

    // Naziv grupe (prva stavka)
    const prvaStavka = STRUKE_LIST.find(s => s.grupa === grupa);
    const grupaNaziv = prvaStavka ? prvaStavka.naziv.replace(/\/\d+$/, "").replace(/\d+$/, "").trim() : "";
    
    // Nađi zajednički naziv grupe
    const stavkeGrupe = STRUKE_LIST.filter(s => s.grupa === grupa);
    let grupaNaslov = stavkeGrupe.length === 1 ? stavkeGrupe[0].naziv : extractGrupaNaslov(stavkeGrupe);
    
    groupHeader.innerHTML = `
      <span style="color:#6b7280;font-size:11px;min-width:32px;font-family:monospace;">${grupa}</span>
      <span style="flex:1;">${grupaNaslov}</span>
    `;
    
    // Dugme za dodavanje custom stavke
    const addBtn = document.createElement("button");
    addBtn.textContent = "+ Dodaj svesku";
    addBtn.style.cssText = `
      display:block; margin:6px 12px 8px; padding:4px 10px;
      background:none; border:1px dashed #93c5fd; border-radius:5px;
      color:#1d4ed8; font-size:12px; cursor:pointer;
    `;
    addBtn.addEventListener("click", () => addCustomStavka(grupa, groupBody, addBtn));

    groupDiv.appendChild(groupHeader);
    groupDiv.appendChild(groupBody);
    groupDiv.appendChild(addBtn);
    body.appendChild(groupDiv);
  }

  setPickerFooter([
    { label: "Otkaži", onClick: closeGitHubTemplateModal },
    { label: "Otvori dokument →", primary: true, onClick: () => {
      const checked = computeOznake();
      console.log("✅ Izabrane struke:", checked);
      // Sačuvaj u global state za kasniju upotrebu pri generisanju tabela
      window._selectedStruke = checked;
      openTemplateFromGitHub();
    }}
  ]);
}

function extractGrupaNaslov(stavke) {
  if (stavke.length === 1) return stavke[0].naziv;
  // Pokušaj da nađeš zajednički deo naziva, inače uzmi naziv prve
  const first = stavke[0].naziv;
  // Pronađi reči koje se pojavljuju u svim nazivima
  const words = first.split(" ");
  let common = "";
  for (let i = words.length; i > 0; i--) {
    const candidate = words.slice(0, i).join(" ");
    if (stavke.every(s => s.naziv.startsWith(candidate))) {
      common = candidate;
      break;
    }
  }
  return common || first;
}

function createStrukaRow(grupa, idx, stavka) {
  const row = document.createElement("div");
  row.style.cssText = `
    display:flex; align-items:center; gap:8px; padding:6px 12px;
    border-top:1px solid #f3f4f6;
  `;
  row.dataset.idx = idx;

  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.checked = stavka.checked;
  cb.style.cssText = "width:15px;height:15px;cursor:pointer;accent-color:#1d4ed8;";
  cb.addEventListener("change", () => {
    _strukeState[grupa][idx].checked = cb.checked;
  });

  const label = document.createElement("span");
  label.style.cssText = "flex:1; font-size:12px; color:#374151;";
  label.textContent = stavka.naziv;

  if (stavka.custom) {
    // Custom stavka ima X dugme
    const delBtn = document.createElement("button");
    delBtn.textContent = "×";
    delBtn.style.cssText = `
      background:none; border:none; color:#9ca3af; font-size:16px;
      cursor:pointer; padding:0 4px; line-height:1;
    `;
    delBtn.addEventListener("click", () => {
      _strukeState[grupa].splice(idx, 1);
      // Re-render grupe
      const groupBody = document.getElementById(`strukaGroup_${grupa}`);
      if (groupBody) {
        groupBody.innerHTML = "";
        _strukeState[grupa].forEach((s, i) => {
          groupBody.appendChild(createStrukaRow(grupa, i, s));
        });
      }
    });
    row.appendChild(cb);
    row.appendChild(label);
    row.appendChild(delBtn);
  } else {
    row.appendChild(cb);
    row.appendChild(label);
  }

  return row;
}

function addCustomStavka(grupa, groupBody, addBtn) {
  // Inline forma za unos naziva
  const formRow = document.createElement("div");
  formRow.style.cssText = "display:flex;gap:6px;padding:6px 12px;border-top:1px solid #f3f4f6;";

  const input = document.createElement("input");
  input.type = "text";
  input.placeholder = "Naziv sveske...";
  input.style.cssText = `
    flex:1; padding:4px 8px; border:1px solid #93c5fd; border-radius:5px;
    font-size:12px; outline:none;
  `;

  const confirmBtn = document.createElement("button");
  confirmBtn.textContent = "Dodaj";
  confirmBtn.style.cssText = `
    padding:4px 10px; background:#1d4ed8; color:#fff; border:none;
    border-radius:5px; font-size:12px; cursor:pointer;
  `;

  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "×";
  cancelBtn.style.cssText = `
    padding:4px 8px; background:#f3f4f6; border:none;
    border-radius:5px; font-size:12px; cursor:pointer;
  `;

  const doAdd = () => {
    const naziv = input.value.trim();
    if (!naziv) { input.focus(); return; }
    _strukeState[grupa].push({ naziv, checked: true, custom: true });
    groupBody.innerHTML = "";
    _strukeState[grupa].forEach((s, i) => {
      groupBody.appendChild(createStrukaRow(grupa, i, s));
    });
    formRow.remove();
  };

  confirmBtn.addEventListener("click", doAdd);
  cancelBtn.addEventListener("click", () => formRow.remove());
  input.addEventListener("keydown", e => { if (e.key === "Enter") doAdd(); if (e.key === "Escape") formRow.remove(); });

  formRow.appendChild(input);
  formRow.appendChild(confirmBtn);
  formRow.appendChild(cancelBtn);
  groupBody.appendChild(formRow);
  setTimeout(() => input.focus(), 50);
}

// Otvori templejt iz GitHub-a
async function openTemplateFromGitHub() {
  const body = el("githubTemplateBody");
  const fileName = _pickerSelectedFile.name;
  const branchId = _pickerSelectedBranch.id;

  if (body) {
    body.innerHTML = `
      <div style="text-align:center;padding:32px;color:#374151;">
        <div style="font-size:36px;margin-bottom:12px;">⏳</div>
        <div style="font-weight:600;font-size:14px;">Preuzimam templejt...</div>
        <div style="font-size:12px;color:#6b7280;margin-top:6px;">${fileName}</div>
      </div>
    `;
  }
  el("githubTemplateFooter").innerHTML = "";

  try {
    const url = buildGitHubRawUrl(branchId, fileName);
    console.log("📥 Skidamo:", url);

    const arrayBuffer = await downloadFileContent(url);

    await Word.run(async (context) => {
      const uint8Array = new Uint8Array(arrayBuffer);
      let binary = "";
      for (let i = 0; i < uint8Array.length; i++) {
        binary += String.fromCharCode(uint8Array[i]);
      }
      const base64 = btoa(binary);
      const doc = context.application.createDocument(base64);
      doc.open();
      await context.sync();
    });

    closeGitHubTemplateModal();
    setStatus(`Otvoren: ${fileName}`, "success");

    // Ako je Glavna sveska, otvori TABELE modal da korisnik izabere sveske i generiše tabele
    if (_pickerSelectedFile.isGlavna) {
      setTimeout(() => {
        openTabeleModal();
      }, 1200);
    }

  } catch (error) {
    console.error("❌ Greška:", error);
    setStatus(`Greška: ${error.message}`, "error");
    if (body) {
      body.innerHTML = `
        <div style="text-align:center;padding:24px;color:#dc2626;">
          <div style="font-size:32px;margin-bottom:12px;">❌</div>
          <div style="font-weight:600;">Greška pri preuzimanju</div>
          <div style="font-size:12px;margin-top:6px;">${error.message}</div>
          <button onclick="renderPickerStep1()" style="
            margin-top:14px;padding:8px 16px;background:#1d4ed8;color:#fff;
            border:none;border-radius:6px;cursor:pointer;
          ">Pokušaj ponovo</button>
        </div>
      `;
      setPickerFooter([{ label: "Otkaži", onClick: closeGitHubTemplateModal }]);
    }
  }
}

// ============================================
// POPUNJAVANJE STRUKE TABELA U WORD-U
// ============================================

async function fillStrukeTables(struke) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag");
      await context.sync();

      for (const cc of ccs.items) {
        const tag = (cc.tag || "").trim();

        if (tag === "SVESKE_TABELA") {
          // Tabela: Broj | Naziv
          const table = cc.insertTable(struke.length + 1, 2, Word.InsertLocation.replace, []);
          table.styleBuiltIn = Word.Style.tableGrid;
          
          // Header
          table.getCell(0, 0).value = "Broj sveske";
          table.getCell(0, 1).value = "Naziv";
          
          struke.forEach((s, i) => {
            table.getCell(i + 1, 0).value = s.oznaka;
            table.getCell(i + 1, 1).value = s.naziv;
          });

        } else if (tag === "SVESKE_OPISI") {
          // Podnaslovi: n.n. NAZIV STRUKE
          let text = "";
          struke.forEach((s, i) => {
            if (i > 0) { // Preskočimo 00 Glavna sveska
              text += `${i}. ${s.naziv.toUpperCase()}\n\n`;
            }
          });
          cc.insertText(text.trim(), Word.InsertLocation.replace);
        }
        // SVESKE_PROJEKTANTI se popunjava ručno
      }

      await context.sync();
      setStatus("Tabele struka popunjene!", "success");
    });
  } catch (error) {
    console.error("❌ Greška pri popunjavanju tabela:", error);
    setStatus("Greška pri popunjavanju tabela struka", "error");
  }
}

// ---------- Stare Graph API funkcije (disabled) ----------
// Stare SharePoint funkcije uklonjene - GitHub mode aktivan

// Fallback: Učitaj templejte iz lokalnog XML-a  
async function loadTemplatesFromDocument() {
  try {
    await Word.run(async (context) => {
      const parts = context.document.customXmlParts;
      parts.load("items");
      await context.sync();

      const targetNamespace = "http://biroa.com/word-addin/templates";
      const targetParts = parts.items.filter(
        (p) => p.namespaceUri === targetNamespace
      );

      if (targetParts.length > 0) {
        const xml = targetParts[0].getXml();
        await context.sync();
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml.value, "text/xml");
        const templateNodes = doc.querySelectorAll("template");
        
        templates = [];
        templateNodes.forEach((node) => {
          const id = node.getAttribute("id") || crypto.randomUUID();
          const name = node.getAttribute("name") || "";
          const desc = node.getAttribute("description") || "";
          const fieldsNode = node.querySelector("fields");
          const fields = [];
          
          if (fieldsNode) {
            fieldsNode.querySelectorAll("field").forEach((f) => {
              fields.push({
                field: f.getAttribute("field") || "",
                type: f.getAttribute("type") || "text",
                format: f.getAttribute("format") || "text:auto",
              });
            });
          }
          
          templates.push({ id, name, desc, fields });
        });
        
        console.log("✅ Učitano", templates.length, "templata");
      } else {
        console.log("ℹ️ Nema sačuvanih templata");
        templates = [];
      }
    });
  } catch (err) {
    console.error("Greška pri učitavanju templata:", err);
    templates = [];
  }
}

// Sačuvaj templejte u XML
async function saveTemplatesToDocument() {
  try {
    await Word.run(async (context) => {
      const parts = context.document.customXmlParts;
      parts.load("items");
      await context.sync();

      const targetNamespace = "http://biroa.com/word-addin/templates";
      const targetParts = parts.items.filter(
        (p) => p.namespaceUri === targetNamespace
      );

      targetParts.forEach((p) => p.delete());
      await context.sync();

      let xml = `<?xml version="1.0" encoding="UTF-8"?><root xmlns="${targetNamespace}">`;
      
      templates.forEach((t) => {
        xml += `<template id="${t.id}" name="${escapeXml(t.name)}" description="${escapeXml(t.desc || '')}">`;
        xml += `<fields>`;
        t.fields.forEach((f) => {
          xml += `<field field="${escapeXml(f.field)}" type="${f.type}" format="${f.format}" />`;
        });
        xml += `</fields></template>`;
      });
      
      xml += `</root>`;
      
      parts.add(xml);
      await context.sync();
      
      console.log("✅ Sačuvano", templates.length, "templata");
    });
  } catch (err) {
    console.error("Greška pri čuvanju templata:", err);
    setStatus("Greška pri čuvanju templata", "error");
  }
}

function escapeXml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// Prikaz liste templata
function renderTemplatesList() {
  const list = el("templatesList");
  if (!list) return;
  
  // Filter po pretrazi
  const searchInput = el("templateSearch");
  const query = (searchInput?.value || "").trim().toLowerCase();
  
  const visible = query
    ? templates.filter(t =>
        (t.name || "").toLowerCase().includes(query) ||
        (t.desc || "").toLowerCase().includes(query)
      )
    : templates;
  
  // Prikaz rezultata
  if (visible.length === 0) {
    list.innerHTML = query 
      ? '<div class="empty-state">Nema templata koji odgovaraju pretrazi</div>'
      : '<div class="empty-state">Nema sačuvanih templata</div>';
    return;
  }
  
  list.innerHTML = "";
  
  visible.forEach((t) => {
    const card = document.createElement("div");
    card.className = "template-card";
    
    const header = document.createElement("div");
    header.className = "template-card-header";
    
    const title = document.createElement("h4");
    title.className = "template-card-title";
    title.textContent = t.name;
    header.appendChild(title);
    
    const actions = document.createElement("div");
    actions.className = "template-card-actions";
    
    const btnLoad = document.createElement("button");
    btnLoad.className = "template-card-btn";
    btnLoad.innerHTML = "📥";
    btnLoad.title = "Učitaj templejt";
    btnLoad.addEventListener("click", (e) => {
      e.stopPropagation();
      loadTemplate(t.id);
    });
    
    const btnEdit = document.createElement("button");
    btnEdit.className = "template-card-btn";
    btnEdit.innerHTML = "✏️";
    btnEdit.title = "Izmeni";
    btnEdit.addEventListener("click", (e) => {
      e.stopPropagation();
      openEditTemplateModal(t.id);
    });
    
    const btnDelete = document.createElement("button");
    btnDelete.className = "template-card-btn delete";
    btnDelete.innerHTML = "🗑️";
    btnDelete.title = "Obriši";
    btnDelete.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (confirm(`Da li želiš da obrišeš templejt "${t.name}"?`)) {
        await deleteTemplate(t.id);
      }
    });
    
    actions.appendChild(btnLoad);
    actions.appendChild(btnEdit);
    actions.appendChild(btnDelete);
    header.appendChild(actions);
    card.appendChild(header);
    
    if (t.desc) {
      const desc = document.createElement("div");
      desc.className = "template-card-desc";
      desc.textContent = t.desc;
      card.appendChild(desc);
    }
    
    const meta = document.createElement("div");
    meta.className = "template-card-meta";
    
    const fieldsInfo = document.createElement("span");
    fieldsInfo.className = "template-card-fields";
    fieldsInfo.textContent = `${t.fields.length} ${t.fields.length === 1 ? 'polje' : t.fields.length < 5 ? 'polja' : 'polja'}`;
    meta.appendChild(fieldsInfo);
    
    card.appendChild(meta);
    
    card.addEventListener("click", () => {
      loadTemplate(t.id);
    });
    
    list.appendChild(card);
  });
}

// Učitaj templejt u tabelu
async function loadTemplate(templateId) {
  const template = templates.find((t) => t.id === templateId);
  if (!template) return;
  
  try {
    setStatus(`Učitavam templejt: ${template.name}...`, "info");
    
    // If template is from SharePoint and fields not loaded yet
    if (template.fileId && template.fields.length === 0) {
      console.log("📥 Skidam fajl sa SharePointa:", template.name);
      
      const arrayBuffer = await downloadFileContent(template.fileId);
      template.fields = await extractFieldsFromDocx(arrayBuffer);
      
      if (template.fields.length === 0) {
        setStatus("Templejt nema polja ili nisu pronađena", "error");
        return;
      }
    }
    
    // Load fields into table
    rows = template.fields.map((f) => ({
      id: crypto.randomUUID(),
      field: f.field,
      value: "",
      type: f.type,
      format: f.format,
    }));
    
    renderRows();
    await saveStateToDocument();
    closeTemplatesModal();
    setStatus(`Učitan templejt: ${template.name} (${template.fields.length} polja)`, "success");
  } catch (error) {
    console.error("❌ Greška pri učitavanju templata:", error);
    setStatus("Greška pri učitavanju templata", "error");
  }
}

// Obriši templejt
async function deleteTemplate(templateId) {
  templates = templates.filter((t) => t.id !== templateId);
  await saveTemplatesToDocument();
  renderTemplatesList();
  setStatus("Templejt obrisan", "success");
}

// Otvori modal za templejte
function openTemplatesModal() {
  const backdrop = el("modalTemplatesBackdrop");
  const modal = el("modalTemplates");
  if (backdrop) backdrop.classList.remove("hidden");
  if (modal) modal.classList.remove("hidden");
  
  // Opciono: osveži svaki put (ako želiš uvek najnovije)
  // await loadTemplatesFromSharePoint();
  
  // Resetuj pretragu
  const searchInput = el("templateSearch");
  if (searchInput) searchInput.value = "";
  
  renderTemplatesList();
}

// Zatvori modal za templejte
function closeTemplatesModal() {
  const backdrop = el("modalTemplatesBackdrop");
  const modal = el("modalTemplates");
  if (backdrop) backdrop.classList.add("hidden");
  if (modal) modal.classList.add("hidden");
}

// Otvori modal za editovanje templata
function openEditTemplateModal(templateId = null) {
  editingTemplateId = templateId;
  
  const backdrop = el("modalEditTemplateBackdrop");
  const modal = el("modalEditTemplate");
  const title = el("editTemplateTitle");
  const nameInput = el("templateName");
  const descInput = el("templateDesc");
  const fieldsList = el("templateFieldsList");
  
  if (!backdrop || !modal) return;
  
  if (templateId) {
    // Editing existing
    const template = templates.find((t) => t.id === templateId);
    if (!template) return;
    
    if (title) title.textContent = "Izmeni templejt";
    if (nameInput) nameInput.value = template.name;
    if (descInput) descInput.value = template.desc || "";
    
    // Show template fields
    if (fieldsList) {
      fieldsList.innerHTML = "";
      if (template.fields.length === 0) {
        fieldsList.classList.add("empty");
        fieldsList.textContent = "Nema polja";
      } else {
        fieldsList.classList.remove("empty");
        template.fields.forEach((f) => {
          const tag = document.createElement("span");
          tag.className = "template-field-tag";
          tag.textContent = f.field;
          fieldsList.appendChild(tag);
        });
      }
    }
  } else {
    // New template
    if (title) title.textContent = "Novi templejt";
    if (nameInput) nameInput.value = "";
    if (descInput) descInput.value = "";
    
    // Show current table fields
    if (fieldsList) {
      fieldsList.innerHTML = "";
      if (rows.length === 0) {
        fieldsList.classList.add("empty");
        fieldsList.textContent = "Nema polja u tabeli";
      } else {
        fieldsList.classList.remove("empty");
        rows.forEach((r) => {
          if (r.field) {
            const tag = document.createElement("span");
            tag.className = "template-field-tag";
            tag.textContent = r.field;
            fieldsList.appendChild(tag);
          }
        });
      }
    }
  }
  
  backdrop.classList.remove("hidden");
  modal.classList.remove("hidden");
  
  if (nameInput) nameInput.focus();
}

// Zatvori modal za editovanje templata
function closeEditTemplateModal() {
  const backdrop = el("modalEditTemplateBackdrop");
  const modal = el("modalEditTemplate");
  if (backdrop) backdrop.classList.add("hidden");
  if (modal) modal.classList.add("hidden");
  editingTemplateId = null;
}

// Sačuvaj templejt
async function saveTemplate() {
  const nameInput = el("templateName");
  const descInput = el("templateDesc");
  
  if (!nameInput) return;
  
  const name = nameInput.value.trim();
  if (!name) {
    setStatus("Unesi ime templata", "error");
    nameInput.focus();
    return;
  }
  
  const desc = descInput ? descInput.value.trim() : "";
  
  if (editingTemplateId) {
    // Update existing
    const template = templates.find((t) => t.id === editingTemplateId);
    if (template) {
      template.name = name;
      template.desc = desc;
      // Keep existing fields
    }
  } else {
    // Create new from current table
    const fields = rows
      .filter((r) => r.field.trim())
      .map((r) => ({
        field: r.field,
        type: r.type,
        format: r.format,
      }));
    
    if (fields.length === 0) {
      setStatus("Nema polja u tabeli za čuvanje", "error");
      return;
    }
    
    templates.push({
      id: crypto.randomUUID(),
      name,
      desc,
      fields,
    });
  }
  
  await saveTemplatesToDocument();
  closeEditTemplateModal();
  setStatus(`Templejt "${name}" sačuvan`, "success");
  
  // If templates modal is open, refresh it
  const templatesModal = el("modalTemplates");
  if (templatesModal && !templatesModal.classList.contains("hidden")) {
    renderTemplatesList();
  }
}

// ---------- wiring ----------
function bindUi() {
  const btnInsert = el("btnInsert");
  const btnFill = el("btnFill");
  const btnClear = el("btnClear");
  const btnDelete = el("btnDelete");
  const btnTemplates = el("btnTemplates");
  const btnAddRow = el("btnAddRow");
  const btnExportCSV = el("btnExportCSV");
  const btnImportCSV = el("btnImportCSV");

  const btnModalClose = el("btnModalClose");
  const btnModalCancel = el("btnModalCancel");
  const btnModalOk = el("btnModalOk");
  const modalBackdrop = el("modalBackdrop");

  // Delete modal buttons
  const btnDeleteModalClose = el("btnDeleteModalClose");
  const btnDeleteCancel = el("btnDeleteCancel");
  const btnDeleteConfirm = el("btnDeleteConfirm");

  // Templates modal buttons
  const btnModalTemplatesClose = el("btnModalTemplatesClose");
  const btnNewTemplate = el("btnNewTemplate");
  const btnModalEditTemplateClose = el("btnModalEditTemplateClose");
  const btnCancelEditTemplate = el("btnCancelEditTemplate");
  const btnSaveTemplate = el("btnSaveTemplate");

  if (btnInsert) btnInsert.addEventListener("click", insertFieldAtSelection);
  if (btnFill) btnFill.addEventListener("click", fillFieldsFromTable);
  if (btnClear) btnClear.addEventListener("click", clearFieldsKeepControls);
  if (btnDelete) btnDelete.addEventListener("click", deleteControlsAndXml);
  if (btnTemplates) btnTemplates.addEventListener("click", openGitHubTemplateModal);
  const btnTables = el("btnTables");
  if (btnTables) btnTables.addEventListener("click", openTabeleModal);
  if (btnExportCSV) btnExportCSV.addEventListener("click", exportCSV);
  if (btnImportCSV) btnImportCSV.addEventListener("click", importCSV);

  if (btnAddRow) {
    btnAddRow.addEventListener("click", () => {
      rows.push({ 
        id: crypto.randomUUID(), 
        field: "", 
        value: "", 
        type: "text", 
        format: "text:auto" 
      });
      renderRows();
      saveStateToDocument();
    });
  }

  if (btnModalClose) btnModalClose.addEventListener("click", closeModal);
  if (btnModalCancel) btnModalCancel.addEventListener("click", closeModal);
  if (btnModalOk) btnModalOk.addEventListener("click", saveModalChanges);
  
  // Delete modal events
  if (btnDeleteModalClose) btnDeleteModalClose.addEventListener("click", closeDeleteModal);
  if (btnDeleteCancel) btnDeleteCancel.addEventListener("click", closeDeleteModal);
  if (btnDeleteConfirm) {
    btnDeleteConfirm.addEventListener("click", async () => {
      closeDeleteModal();
      await performDelete();
    });
  }
  
  // Templates modal events
  if (btnModalTemplatesClose) btnModalTemplatesClose.addEventListener("click", closeTemplatesModal);
  if (btnNewTemplate) btnNewTemplate.addEventListener("click", () => openEditTemplateModal(null));
  
  // Templates toolbar events
  const btnRefreshTemplates = el("btnRefreshTemplates");
  const templateSearch = el("templateSearch");
  
  if (btnRefreshTemplates) {
    btnRefreshTemplates.addEventListener("click", async () => {
      setStatus("Osvežavam templejte...", "info");
      await loadTemplatesFromSharePoint();
      renderTemplatesList();
    });
  }
  
  if (templateSearch) {
    templateSearch.addEventListener("input", () => {
      renderTemplatesList();
    });
  }
  
  // Edit template modal events
  if (btnModalEditTemplateClose) btnModalEditTemplateClose.addEventListener("click", closeEditTemplateModal);
  if (btnCancelEditTemplate) btnCancelEditTemplate.addEventListener("click", closeEditTemplateModal);
  if (btnSaveTemplate) btnSaveTemplate.addEventListener("click", saveTemplate);
  
  // Spreci da klik NA modal zatvara modal
  const modal = el("modal");
  const deleteModal = el("deleteModal");
  const templatesModal = el("modalTemplates");
  const editTemplateModal = el("modalEditTemplate");
  if (modal) modal.addEventListener("click", (e) => e.stopPropagation());
  if (deleteModal) deleteModal.addEventListener("click", (e) => e.stopPropagation());
  if (templatesModal) templatesModal.addEventListener("click", (e) => e.stopPropagation());
  if (editTemplateModal) editTemplateModal.addEventListener("click", (e) => e.stopPropagation());
  
  // Backdrop zatvara modal SAMO ako klikneš na backdrop (ne na modal)
  if (modalBackdrop) {
    modalBackdrop.addEventListener("click", (e) => {
      if (e.target !== modalBackdrop) return; // samo klik na "prazno"
      closeModal();
      closeDeleteModal();
    });
  }
  
  // Templates backdrop
  const templatesBackdrop = el("modalTemplatesBackdrop");
  if (templatesBackdrop) {
    templatesBackdrop.addEventListener("click", (e) => {
      if (e.target !== templatesBackdrop) return;
      closeTemplatesModal();
    });
  }
  
  // Edit template backdrop
  const editTemplateBackdrop = el("modalEditTemplateBackdrop");
  if (editTemplateBackdrop) {
    editTemplateBackdrop.addEventListener("click", (e) => {
      if (e.target !== editTemplateBackdrop) return;
      closeEditTemplateModal();
    });
  }
}

Office.onReady(async () => {
  console.log("✅ Office.onReady STARTED");
  
  try {
    console.log("🔄 Pozivam loadStateFromDocument...");
    await loadStateFromDocument();
    console.log("✅ loadStateFromDocument završen, rows.length:", rows.length);
    
    console.log("🔄 Pozivam loadTemplatesFromSharePoint...");
    
    // Dodaj timeout da ne blokira renderRows ako se SharePoint zaglavi
    try {
      await Promise.race([
        loadTemplatesFromSharePoint(),
        new Promise((_, reject) => 
          setTimeout(() => reject(new Error("SharePoint timeout")), 5000)
        )
      ]);
      console.log("✅ loadTemplatesFromSharePoint završen");
    } catch (timeoutError) {
      console.warn("⚠️ loadTemplatesFromSharePoint timeout ili error:", timeoutError.message);
      console.log("   Nastavljam dalje sa renderovanjem...");
    }
    
  } catch (e) {
    console.error("❌ Load state error:", e);
  }

  console.log("🎨 Pozivam renderRows sa rows.length:", rows.length);
  renderRows();
  console.log("✅ renderRows završen");
  
  console.log("🔗 Pozivam bindUi...");
  bindUi();
  console.log("✅ bindUi završen");
  
  console.log("✅✅✅ Office.onReady COMPLETED ✅✅✅");
});

// ============================================
// TABELE MODULE (V58)
// Generisanje tabela u Glavnoj svesci
// ============================================

// Definicija svih struka sa prefiksom za naziv u dokumentu
const TABELE_STRUKE = [
  // Grupa 01 - Arhitektura
  { oznaka: "01",   naziv: "ARHITEKTURA",                                    grupa: "01", grupaNaziv: "Arhitektura",             tip: "projekat" },
  // Grupa 02 - Konstrukcija
  { oznaka: "02",   naziv: "KONSTRUKCIJA",                                   grupa: "02", grupaNaziv: "Konstrukcija",            tip: "projekat" },
  { oznaka: "02",   naziv: "SAOBRAĆAJNA KONSTRUKCIJA",                       grupa: "02", grupaNaziv: "Konstrukcija",            tip: "projekat" },
  // Grupa 03 - Hidro
  { oznaka: "03",   naziv: "HIDROTEHNIČKE INSTALACIJE",                      grupa: "03", grupaNaziv: "Hidrotehničke inst.",      tip: "projekat" },
  // Grupa 04 - Elektro
  { oznaka: "04",   naziv: "ELEKTROENERGETSKE INSTALACIJE",                  grupa: "04", grupaNaziv: "Elektroenergetika",       tip: "projekat" },
  { oznaka: "04",   naziv: "TRAFO STANICA",                                  grupa: "04", grupaNaziv: "Elektroenergetika",       tip: "projekat" },
  // Grupa 05 - Signalne
  { oznaka: "05",   naziv: "SIGNALNE INSTALACIJE",                           grupa: "05", grupaNaziv: "Signalne inst.",          tip: "projekat" },
  { oznaka: "05",   naziv: "STABILNI SISTEM ZA DOJAVU POŽARA",               grupa: "05", grupaNaziv: "Signalne inst.",          tip: "projekat" },
  { oznaka: "05",   naziv: "VIDEO NADZOR",                                   grupa: "05", grupaNaziv: "Signalne inst.",          tip: "projekat" },
  // Grupa 06 - Mašinske
  { oznaka: "06",   naziv: "TERMOTEHNIČKE INSTALACIJE",                      grupa: "06", grupaNaziv: "Mašinske inst.",          tip: "projekat" },
  { oznaka: "06",   naziv: "PUTNIČKI LIFT",                                  grupa: "06", grupaNaziv: "Mašinske inst.",          tip: "projekat" },
  { oznaka: "06",   naziv: "AUTOLIFT",                                       grupa: "06", grupaNaziv: "Mašinske inst.",          tip: "projekat" },
  { oznaka: "06",   naziv: "STABILNI SISTEM ZA GAŠENJE POŽARA",              grupa: "06", grupaNaziv: "Mašinske inst.",          tip: "projekat" },
  { oznaka: "06",   naziv: "VENTILACIJA I NADPRITISAK GARAŽE",               grupa: "06", grupaNaziv: "Mašinske inst.",          tip: "projekat" },
  // Grupa 07
  { oznaka: "07",   naziv: "TEHNOLOGIJA",                                    grupa: "07", grupaNaziv: "Tehnologija",             tip: "projekat" },
  // Grupa 08
  { oznaka: "08",   naziv: "SAOBRAĆAJ I SAOBRAĆAJNA SIGNALIZACIJA",         grupa: "08", grupaNaziv: "Saobraćaj",              tip: "projekat" },
  // Grupa 09
  { oznaka: "09",   naziv: "SPOLJNO UREĐENJE",                               grupa: "09", grupaNaziv: "Spoljno uređenje",       tip: "projekat" },
  // Grupa 10
  { oznaka: "10",   naziv: "PRIPREMNI RADOVI - PROJEKAT RUŠENJA",            grupa: "10", grupaNaziv: "Pripremni radovi",       tip: "projekat" },
  { oznaka: "10",   naziv: "PRIPREMNI RADOVI - PROJEKAT OBEZBEĐENJA TEMELJNE JAME", grupa: "10", grupaNaziv: "Pripremni radovi", tip: "projekat" },
  // Elaborati (posebna grupa - nemaju podbroj, uvek singularni)
  { oznaka: "EE",   naziv: "ELABORAT ENERGETSKE EFIKASNOSTI",                grupa: "EE",   grupaNaziv: "Elaborat EE",          tip: "elaborat" },
  { oznaka: "GEO",  naziv: "ELABORAT O GEOTEHNIČKIM USLOVIMA IZRADE PGD",   grupa: "GEO",  grupaNaziv: "Elaborat GEO",         tip: "elaborat" },
  { oznaka: "EZOP", naziv: "ELABORAT ZAŠTITE OD POŽARA",                    grupa: "EZOP", grupaNaziv: "Elaborat EZOP",        tip: "elaborat" },
];

// Oznaka 0.8 NASLOVI - mapiranje oznake -> naziv poglavlja opisnog teksta
const TABELE_08_NAZIVI = {
  "01":   "ARHITEKTONSKI OPIS",
  "02":   "OPIS KONSTRUKCIJE",
  "03":   "OPIS HIDROTEHNIČKIH INSTALACIJA",
  "04":   "OPIS ELEKTROENERGETSKIH INSTALACIJA",
  "05":   "OPIS SIGNALNIH INSTALACIJA",
  "06":   "OPIS MAŠINSKIH INSTALACIJA",
  "07":   "OPIS TEHNOLOGIJE",
  "08":   "OPIS SAOBRAĆAJA I SAOBRAĆAJNE SIGNALIZACIJE",
  "09":   "OPIS SPOLJNOG UREĐENJA",
  "10":   "OPIS PRIPREMNIH RADOVA",
  "EE":   "ELABORAT ENERGETSKE EFIKASNOSTI",
  "GEO":  "ELABORAT O GEOTEHNIČKIM USLOVIMA",
  "EZOP": "ELABORAT ZAŠTITE OD POŽARA",
};

// State za tabele modal
let _tabeleStrukeState = {}; 
// Format: { grupa: [ { naziv, checked, custom, id } ] }

// Computed rezultat - lista izabranih svezaka sa konačnim oznakama
// [ { oznaka: "06.1", naziv: "TERMOTEHNIČKE INSTALACIJE", tip: "projekat"|"elaborat" } ]
let _tabeleSveske = [];

// ---- Inicijalizacija state-a ----
function initTabeleState() {
  _tabeleStrukeState = {};

  const grupe = {};
  for (const s of TABELE_STRUKE) {
    if (!grupe[s.grupa]) grupe[s.grupa] = [];
    grupe[s.grupa].push({
      id: crypto.randomUUID(),
      naziv: s.naziv,
      checked: false,
      custom: false,
      tip: s.tip,
      grupaNaziv: s.grupaNaziv,
    });
  }
  _tabeleStrukeState = grupe;
}

// ---- Izračunaj finalne oznake ----
function computeTabeleOznake() {
  const result = [];

  // Uvek prva: Glavna sveska
  result.push({ oznaka: "0", naziv: "GLAVNA SVESKA", tip: "glavna" });

  for (const [grupa, stavke] of Object.entries(_tabeleStrukeState)) {
    const checked = stavke.filter(s => s.checked);
    if (checked.length === 0) continue;

    const isElaborat = checked[0].tip === "elaborat";

    if (isElaborat || checked.length === 1) {
      // Elaborati i singularni projekti → samo oznaka grupe bez podbrojeva
      result.push({
        oznaka: grupa,
        naziv: checked[0].naziv,
        tip: checked[0].tip,
      });
    } else {
      // Više u grupi → grupa.1, grupa.2...
      checked.forEach((s, i) => {
        result.push({
          oznaka: `${grupa}.${i + 1}`,
          naziv: s.naziv,
          tip: s.tip,
        });
      });
    }
  }

  return result;
}

// ---- Tag za placeholder tabele u Word-u ----
// Format: BA_TABLE|type=04   (za tabelu 0.4)
//         BA_TABLE|type=05
//         BA_TABLE|type=061
//         BA_TABLE|type=062
//         BA_TABLE|type=08
function makeTableTag(type) {
  return `BA_TABLE|type=${type}`;
}

// ============================================
// MODAL - Otvaranje i zatvaranje
// ============================================

function openTabeleModal() {
  // Kreira backdrop + modal dinamički ako ne postoje
  let backdrop = el("tabeleBackdrop");
  if (!backdrop) {
    backdrop = document.createElement("div");
    backdrop.id = "tabeleBackdrop";
    backdrop.style.cssText = `
      position:fixed; inset:0; background:rgba(0,0,0,0.5);
      z-index:1000; display:flex; align-items:center; justify-content:center;
    `;
    backdrop.addEventListener("click", e => {
      if (e.target === backdrop) closeTabeleModal();
    });

    const modal = document.createElement("div");
    modal.id = "tabeleModal";
    modal.style.cssText = `
      background:#fff; border-radius:12px; width:520px; max-width:97vw;
      max-height:92vh; display:flex; flex-direction:column;
      box-shadow:0 8px 32px rgba(0,0,0,0.2);
      font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
    `;

    const header = document.createElement("div");
    header.id = "tabeleHeader";
    header.style.cssText = `
      background:#1d4ed8; color:#fff; padding:14px 18px;
      display:flex; align-items:center; justify-content:space-between;
      border-radius:12px 12px 0 0; flex-shrink:0;
    `;

    const body = document.createElement("div");
    body.id = "tabeleBody";
    body.style.cssText = "padding:18px; overflow-y:auto; flex:1; min-height:0;";

    const footer = document.createElement("div");
    footer.id = "tabeleFooter";
    footer.style.cssText = `
      padding:12px 18px; border-top:1px solid #e5e7eb;
      display:flex; gap:8px; justify-content:flex-end; flex-shrink:0;
    `;

    modal.appendChild(header);
    modal.appendChild(body);
    modal.appendChild(footer);
    backdrop.appendChild(modal);
    document.body.appendChild(backdrop);
  }

  initTabeleState();
  renderTabeleStep1();
  el("tabeleBackdrop").style.display = "flex";
}

function closeTabeleModal() {
  const backdrop = el("tabeleBackdrop");
  if (backdrop) backdrop.style.display = "none";
}

// ---- Header / Footer helpers ----
function setTabeleHeader(title, showBack, onBack) {
  const header = el("tabeleHeader");
  if (!header) return;
  header.innerHTML = `
    <div style="display:flex; align-items:center; gap:10px;">
      ${showBack
        ? `<button id="tabeleBackBtn" style="background:rgba(255,255,255,0.2);border:none;color:#fff;
            font-size:16px;cursor:pointer;border-radius:4px;padding:2px 8px;">←</button>`
        : ""}
      <span style="font-weight:700; font-size:15px;">🗂️ ${title}</span>
    </div>
    <button id="tabeleCloseBtn" style="background:none;border:none;color:#fff;font-size:22px;cursor:pointer;line-height:1;">×</button>
  `;
  el("tabeleCloseBtn").addEventListener("click", closeTabeleModal);
  if (showBack) el("tabeleBackBtn").addEventListener("click", onBack);
}

function setTabeleFooter(buttons) {
  const footer = el("tabeleFooter");
  if (!footer) return;
  footer.innerHTML = "";
  buttons.forEach(b => {
    const btn = document.createElement("button");
    btn.textContent = b.label;
    btn.disabled = b.disabled || false;
    btn.style.cssText = b.primary
      ? `padding:9px 22px;background:#1d4ed8;color:#fff;border:none;border-radius:7px;
         cursor:pointer;font-size:13px;font-weight:600;transition:all 0.15s;`
      : `padding:9px 18px;border:1px solid #d1d5db;border-radius:7px;background:#fff;
         cursor:pointer;font-size:13px;color:#374151;transition:all 0.15s;`;
    if (b.disabled) btn.style.opacity = "0.5";
    btn.addEventListener("click", b.onClick);
    if (b.id) btn.id = b.id;
    footer.appendChild(btn);
  });
}

// ============================================
// KORAK 1 - Izbor svezaka
// ============================================

function renderTabeleStep1() {
  setTabeleHeader("Izaberi sveske projekta", false, null);
  const body = el("tabeleBody");
  body.innerHTML = "";

  // Info box
  const info = document.createElement("div");
  info.className = "info-box";
  info.innerHTML = `
    Izaberi koje sveske (projekti/elaborati) postoje u ovom projektu.<br>
    <strong>Glavna sveska (0)</strong> se uvek uključuje automatski.<br>
    Grupe sa više izabranih stavki dobijaju podbroj: <strong>06.1, 06.2...</strong>
  `;
  body.appendChild(info);

  // Grupe svezaka
  const grid = document.createElement("div");
  grid.className = "sveske-grid";

  const grupeRedosled = ["01","02","03","04","05","06","07","08","09","10","EE","GEO","EZOP"];

  for (const grupa of grupeRedosled) {
    const stavke = _tabeleStrukeState[grupa];
    if (!stavke || stavke.length === 0) continue;

    const protoStavka = TABELE_STRUKE.find(s => s.grupa === grupa);
    const grupaNaziv = protoStavka ? protoStavka.grupaNaziv : grupa;
    const tipGrupe = protoStavka ? protoStavka.tip : "projekat";

    const block = document.createElement("div");
    block.className = "struka-group-block";
    block.id = `tabGrupa_${grupa}`;

    const ghdr = document.createElement("div");
    ghdr.className = "struka-group-header";
    ghdr.innerHTML = `
      <span class="oznaka-badge">${grupa}</span>
      <span>${grupaNaziv}</span>
      ${tipGrupe === "elaborat"
        ? `<span style="margin-left:auto;background:#fef3c7;color:#92400e;padding:1px 7px;
             border-radius:4px;font-size:10px;font-weight:600;">elaborat</span>`
        : ""}
    `;
    block.appendChild(ghdr);

    const gbody = document.createElement("div");
    gbody.id = `tabGrupaBody_${grupa}`;

    stavke.forEach((s, idx) => {
      gbody.appendChild(makeStrukaCheckRow(grupa, idx, s));
    });

    block.appendChild(gbody);

    // Dugme za dodavanje custom sveske (samo za projekte, ne elaborate)
    if (tipGrupe === "projekat") {
      const addWrap = document.createElement("div");
      addWrap.className = "add-custom-row";
      const addBtn = document.createElement("button");
      addBtn.className = "add-custom-btn";
      addBtn.textContent = "+ Dodaj svesku u grupu";
      addBtn.addEventListener("click", () => {
        showAddCustomForm(grupa, gbody, addWrap);
      });
      addWrap.appendChild(addBtn);
      block.appendChild(addWrap);
    }

    grid.appendChild(block);
  }

  body.appendChild(grid);

  setTabeleFooter([
    { label: "Otkaži", onClick: closeTabeleModal },
    {
      label: "Sledeće: Pregled →",
      primary: true,
      id: "tabeleNextBtn",
      onClick: () => {
        _tabeleSveske = computeTabeleOznake();
        renderTabeleStep2();
      }
    }
  ]);
}

function makeStrukaCheckRow(grupa, idx, stavka) {
  const row = document.createElement("div");
  row.className = "struka-row";
  row.id = `tabRow_${stavka.id}`;

  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.id = `tabCb_${stavka.id}`;
  cb.checked = stavka.checked;
  cb.addEventListener("change", () => {
    _tabeleStrukeState[grupa][idx].checked = cb.checked;
  });

  const lbl = document.createElement("label");
  lbl.htmlFor = `tabCb_${stavka.id}`;
  lbl.innerHTML = `<span class="naziv-tekst">${stavka.naziv}</span>`;
  if (stavka.custom) {
    lbl.style.fontStyle = "italic";
  }

  if (stavka.custom) {
    const delBtn = document.createElement("button");
    delBtn.textContent = "×";
    delBtn.title = "Ukloni";
    delBtn.style.cssText = `background:none;border:none;color:#9ca3af;font-size:16px;cursor:pointer;padding:0 4px;line-height:1;`;
    delBtn.addEventListener("click", () => {
      _tabeleStrukeState[grupa].splice(idx, 1);
      reRenderGrupaBody(grupa);
    });
    row.appendChild(cb);
    row.appendChild(lbl);
    row.appendChild(delBtn);
  } else {
    row.appendChild(cb);
    row.appendChild(lbl);
  }

  return row;
}

function reRenderGrupaBody(grupa) {
  const gbody = el(`tabGrupaBody_${grupa}`);
  if (!gbody) return;
  gbody.innerHTML = "";
  (_tabeleStrukeState[grupa] || []).forEach((s, i) => {
    gbody.appendChild(makeStrukaCheckRow(grupa, i, s));
  });
}

function showAddCustomForm(grupa, gbody, addWrap) {
  // Sakrij dugme privremeno
  addWrap.style.display = "none";

  const formRow = document.createElement("div");
  formRow.style.cssText = "display:flex;gap:6px;padding:5px 12px 8px;";

  const inp = document.createElement("input");
  inp.type = "text";
  inp.placeholder = "Naziv sveske (npr. HIDRANTSKA MREŽA)...";
  inp.style.cssText = `flex:1;padding:5px 8px;border:1px solid #93c5fd;border-radius:5px;
    font-size:12px;outline:none;font-family:inherit;`;

  const okBtn = document.createElement("button");
  okBtn.textContent = "Dodaj";
  okBtn.style.cssText = `padding:5px 12px;background:#1d4ed8;color:#fff;border:none;
    border-radius:5px;font-size:12px;cursor:pointer;`;

  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "×";
  cancelBtn.style.cssText = `padding:5px 9px;background:#f3f4f6;border:none;
    border-radius:5px;font-size:12px;cursor:pointer;`;

  const doAdd = () => {
    const naziv = inp.value.trim().toUpperCase();
    if (!naziv) { inp.focus(); return; }
    _tabeleStrukeState[grupa].push({
      id: crypto.randomUUID(),
      naziv,
      checked: true,
      custom: true,
      tip: "projekat",
    });
    reRenderGrupaBody(grupa);
    formRow.remove();
    addWrap.style.display = "";
  };

  okBtn.addEventListener("click", doAdd);
  cancelBtn.addEventListener("click", () => {
    formRow.remove();
    addWrap.style.display = "";
  });
  inp.addEventListener("keydown", e => {
    if (e.key === "Enter") doAdd();
    if (e.key === "Escape") { formRow.remove(); addWrap.style.display = ""; }
  });

  formRow.appendChild(inp);
  formRow.appendChild(okBtn);
  formRow.appendChild(cancelBtn);
  gbody.parentElement.insertBefore(formRow, addWrap);
  setTimeout(() => inp.focus(), 50);
}

// ============================================
// KORAK 2 - Pregled i potvrda
// ============================================

function renderTabeleStep2() {
  setTabeleHeader("Pregled tabela", true, renderTabeleStep1);
  const body = el("tabeleBody");
  body.innerHTML = "";

  // Filtriraj samo izabrane (bez glavne sveske koja je uvek tu)
  const izabrane = _tabeleSveske.filter(s => s.tip !== "glavna");
  const projekti = _tabeleSveske.filter(s => s.tip === "projekat");
  const elaborati = _tabeleSveske.filter(s => s.tip === "elaborat");

  if (izabrane.length === 0) {
    body.innerHTML = `
      <div style="text-align:center;padding:40px 20px;color:#9ca3af;">
        <div style="font-size:40px;margin-bottom:10px;">⚠️</div>
        <div style="font-weight:600;color:#374151;">Nijedna sveska nije izabrana</div>
        <div style="font-size:12px;margin-top:6px;">Vrati se i izaberi bar jednu svesku projekta.</div>
      </div>`;
    setTabeleFooter([
      { label: "← Nazad", onClick: renderTabeleStep1 },
      { label: "Otkaži", onClick: closeTabeleModal }
    ]);
    return;
  }

  // Info
  const info = document.createElement("div");
  info.className = "info-box";
  info.innerHTML = `
    Izabrano je <strong>${izabrane.length}</strong> svezaka.<br>
    Klikni <strong>Generiši tabele</strong> da se Word placeholder tabele popune redovima.<br>
    <span style="color:#6b7280;font-size:11px;">
      Placeholder tabele moraju imati tag: <code>BA_TABLE|type=04</code> itd.
    </span>
  `;
  body.appendChild(info);

  // Preview tabele 0.4 i 0.5
  body.appendChild(makePreviewTabela04_05(_tabeleSveske));

  // Preview tabele 0.6.1 (samo projekti, bez elaborata)
  if (projekti.length > 0) {
    body.appendChild(makePreviewTabela061(projekti));
  }

  // Preview tabele 0.6.2 (samo elaborati)
  if (elaborati.length > 0) {
    body.appendChild(makePreviewTabela062(elaborati));
  }

  // Preview 0.8 naslovi
  body.appendChild(makePreview08(_tabeleSveske));

  setTabeleFooter([
    { label: "← Nazad", onClick: renderTabeleStep1 },
    { label: "Otkaži", onClick: closeTabeleModal },
    {
      label: "✅ Generiši tabele u Word-u",
      primary: true,
      id: "tabeleGenerisiBtn",
      onClick: () => generisiTabele()
    }
  ]);
}

function makePreviewTabela04_05(sveske) {
  const wrap = document.createElement("div");
  wrap.style.marginBottom = "14px";

  const title = document.createElement("div");
  title.style.cssText = "font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;";
  title.textContent = "Tabele 0.4 i 0.5 (Izjava + Sadržaj tehničke dokumentacije)";
  wrap.appendChild(title);

  const table = document.createElement("div");
  table.className = "table-preview";

  const hdr = document.createElement("div");
  hdr.className = "table-preview-header";
  hdr.textContent = "0.4 / 0.5 — Br. sveske | Naziv | Br. licence (prazno)";
  table.appendChild(hdr);

  const tbl = document.createElement("table");
  const thead = tbl.insertRow();
  thead.style.background = "#f3f4f6";
  [["Br.", "40px"], ["Naziv sveske", "auto"], ["Br. licence", "90px"]].forEach(([t, w]) => {
    const th = document.createElement("th");
    th.textContent = t;
    th.style.width = w;
    thead.appendChild(th);
  });

  sveske.forEach(s => {
    const tr = tbl.insertRow();
    const td1 = tr.insertCell(); td1.textContent = s.oznaka; td1.className = "auto-col";
    const nazivTekst = s.tip === "glavna" ? "GLAVNA SVESKA" : `PROJEKAT ${s.naziv}`;
    const td2 = tr.insertCell(); td2.textContent = nazivTekst; td2.className = "auto-col";
    const td3 = tr.insertCell(); td3.textContent = "..."; td3.className = "empty-col";
  });

  table.appendChild(tbl);
  wrap.appendChild(table);
  return wrap;
}

function makePreviewTabela061(projekti) {
  const wrap = document.createElement("div");
  wrap.style.marginBottom = "14px";

  const title = document.createElement("div");
  title.style.cssText = "font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;";
  title.textContent = "Tabela 0.6.1 (Podaci o projektantima)";
  wrap.appendChild(title);

  const table = document.createElement("div");
  table.className = "table-preview";

  const hdr = document.createElement("div");
  hdr.className = "table-preview-header";
  hdr.textContent = "0.6.1 — Leva kolona (auto) | Desna kolona (prazno za unos)";
  table.appendChild(hdr);

  const tbl = document.createElement("table");
  const thr = tbl.insertRow();
  thr.style.background = "#f3f4f6";
  const thL = document.createElement("th"); thL.textContent = "Leva kolona (auto)"; thr.appendChild(thL);
  const thR = document.createElement("th"); thR.textContent = "Desna kolona (ručno)"; thr.appendChild(thR);

  projekti.forEach(s => {
    // Blok od 4 reda po svesci
    const blokNaziv = `${s.oznaka}. PROJEKAT ${s.naziv}`;
    const redovi = [
      { l: blokNaziv, r: "" },
      { l: "Projektant:", r: "..." },
      { l: "Broj licence:", r: "..." },
      { l: "Potpis:", r: "" },
    ];
    redovi.forEach(r => {
      const tr = tbl.insertRow();
      const tdL = tr.insertCell(); tdL.className = "auto-col"; tdL.textContent = r.l;
      const tdR = tr.insertCell(); tdR.textContent = r.r; tdR.className = r.r ? "empty-col" : "";
    });
  });

  table.appendChild(tbl);
  wrap.appendChild(table);
  return wrap;
}

function makePreviewTabela062(elaborati) {
  const wrap = document.createElement("div");
  wrap.style.marginBottom = "14px";

  const title = document.createElement("div");
  title.style.cssText = "font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;";
  title.textContent = "Tabela 0.6.2 (Podaci o licima za elaborate)";
  wrap.appendChild(title);

  const table = document.createElement("div");
  table.className = "table-preview";

  const hdr = document.createElement("div");
  hdr.className = "table-preview-header";
  hdr.textContent = "0.6.2 — Elaborati (Izrađivač, Ovlašćeno lice, Br. ovlašćenja, Potpis)";
  table.appendChild(hdr);

  const tbl = document.createElement("table");
  const thr = tbl.insertRow();
  thr.style.background = "#f3f4f6";
  const thL = document.createElement("th"); thL.textContent = "Naziv elaborata + labele"; thr.appendChild(thL);
  const thR = document.createElement("th"); thR.textContent = "Vrednost (ručno)"; thr.appendChild(thR);

  elaborati.forEach(s => {
    const blokNaziv = `${s.naziv}:`;
    [
      { l: blokNaziv, r: "" },
      { l: "Izrađivač:", r: "..." },
      { l: "Ovlašćeno lice:", r: "..." },
      { l: "Broj ovlašćenja:", r: "..." },
      { l: "Potpis:", r: "" },
    ].forEach(r => {
      const tr = tbl.insertRow();
      const tdL = tr.insertCell(); tdL.className = "auto-col"; tdL.textContent = r.l;
      const tdR = tr.insertCell(); tdR.textContent = r.r; tdR.className = r.r ? "empty-col" : "";
    });
  });

  table.appendChild(tbl);
  wrap.appendChild(table);
  return wrap;
}

function makePreview08(sveske) {
  const wrap = document.createElement("div");
  wrap.style.marginBottom = "14px";

  const title = document.createElement("div");
  title.style.cssText = "font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;";
  title.textContent = "Naslovi 0.8 (Sažet tehnički opis)";
  wrap.appendChild(title);

  const table = document.createElement("div");
  table.className = "table-preview";

  const hdr = document.createElement("div");
  hdr.className = "table-preview-header";
  hdr.textContent = "0.8 — Naslovi poglavlja opisnog dela";
  table.appendChild(hdr);

  const tbl = document.createElement("table");
  const thr = tbl.insertRow();
  thr.style.background = "#f3f4f6";
  const th = document.createElement("th"); th.textContent = "Naslov"; thr.appendChild(th);

  // Grupiši - jedan naslov po grupi (ne po podbroju)
  const vidjenePrije = new Set();
  let rbr = 1;
  sveske.forEach(s => {
    if (s.tip === "glavna") return;
    const oznakaGrupe = s.oznaka.includes(".") ? s.oznaka.split(".")[0] : s.oznaka;
    if (vidjenePrije.has(oznakaGrupe)) return;
    vidjenePrije.add(oznakaGrupe);

    const naslovTekst = TABELE_08_NAZIVI[oznakaGrupe] || s.naziv;
    const tr = tbl.insertRow();
    const td = tr.insertCell();
    td.className = "auto-col";
    td.textContent = `0.8.${rbr}. ${naslovTekst}`;
    rbr++;
  });

  table.appendChild(tbl);
  wrap.appendChild(table);
  return wrap;
}

// ============================================
// GENERISANJE TABELA U WORD-U
// ============================================

// ============================================
// TABELE - Hidden Tag Pattern (V63)
// Placeholder: obična Word tabela sa skrivenim ID tagom u prvoj ćeliji
// Tag format: [BA:04], [BA:05], [BA:061], [BA:062], [BA:08]
// Tag je skriven: font.color="white", font.size=1
// ============================================

// ID tagovi za svaki tip tabele
const TABLE_TAGS = {
  "04":  "[BA:04]",
  "05":  "[BA:05]",
  "061": "[BA:061]",
  "062": "[BA:062]",
  "08":  "[BA:08]",
};

// Pronađi tabelu u dokumentu po ID tagu, vrati parentTable ili null
async function nadjiTabeluPoTagu(context, tag) {
  const results = context.document.body.search(tag, { matchCase: true });
  results.load("items");
  await context.sync();
  if (results.items.length === 0) return null;

  const parentTable = results.items[0].parentTable;
  parentTable.load("isNullObject");
  await context.sync();
  if (parentTable.isNullObject) return null;

  return parentTable;
}

// Sakrij ID tag unutar tabele (bela boja, font 1pt)
async function sakrijTag(context, table, tag) {
  const tableRange = table.getRange();
  const found = tableRange.search(tag, { matchCase: true });
  found.load("items");
  await context.sync();
  if (found.items.length > 0) {
    found.items[0].font.color = "white";
    found.items[0].font.size = 1;
    await context.sync();
  }
}

// Zameni staru tabelu novom na istom mestu
// Vraća novu tabelu
async function zameniTabelu(context, staraTabela, noviPodaci, cols) {
  const afterRange = staraTabela.getRange(Word.RangeLocation.after);
  staraTabela.delete();
  await context.sync();

  const novaTabela = afterRange.insertTable(
    noviPodaci.length, cols, Word.InsertLocation.before, noviPodaci
  );
  novaTabela.styleBuiltIn = Word.Style.tableGrid;
  await context.sync();
  return novaTabela;
}

// Upiši podatke u postojeću tabelu (uskladi broj redova, popuni ćelije)
async function popuniTabelu(context, table, podaci, cols) {
  table.load("rowCount");
  await context.sync();

  const targetRows = podaci.length;
  const currentRows = table.rowCount;

  if (currentRows < targetRows) {
    for (let i = currentRows; i < targetRows; i++) {
      table.addRows(Word.InsertLocation.end, 1);
    }
    await context.sync();
  }

  table.load("rows/items");
  await context.sync();

  for (let i = 0; i < targetRows; i++) {
    if (i >= table.rows.items.length) break;
    const row = table.rows.items[i];
    row.load("cells/items");
    await context.sync();

    const cells = row.cells.items;
    for (let c = 0; c < cols; c++) {
      if (!cells[c]) continue;
      cells[c].body.clear();
      const val = (podaci[i] && podaci[i][c] != null) ? String(podaci[i][c]) : "";
      if (val) cells[c].body.insertText(val, Word.InsertLocation.start);
    }
  }

  // Obriši višak redova odozdo
  if (currentRows > targetRows) {
    table.load("rows/items");
    await context.sync();
    for (let i = currentRows - 1; i >= targetRows; i--) {
      if (table.rows.items[i]) table.rows.items[i].delete();
    }
    await context.sync();
  }

  await context.sync();
}

// Poravnaj prvu kolonu desno (da tag ne utiče na izgled prvog reda)
async function poravnajPrvuKolonuDesno(context, table) {
  try {
    table.load("rows/items");
    await context.sync();
    for (const row of table.rows.items) {
      row.load("cells/items");
      await context.sync();
      if (row.cells.items[0]) {
        const paras = row.cells.items[0].body.paragraphs;
        paras.load("items");
        await context.sync();
        if (paras.items.length > 0) {
          paras.items[0].alignment = Word.Alignment.right;
        }
      }
    }
    await context.sync();
  } catch (e) {
    console.warn("⚠️ poravnajPrvuKolonuDesno greška (preskočeno):", e.message);
  }
}

// ---- Glavna funkcija generisanja ----
async function generisiTabele() {
  const btn = el("tabeleGenerisiBtn");
  if (btn) { btn.disabled = true; btn.textContent = "⏳ Generiše se..."; }

  try {
    const sveske   = _tabeleSveske;
    const projekti = sveske.filter(s => s.tip !== "glavna" && s.tip !== "elaborat");
    const elaborati = sveske.filter(s => s.tip === "elaborat");

    let ukupnoTabela = 0;

    await Word.run(async (context) => {

      // ---- TABELA 04 ----
      try {
        const t04 = await nadjiTabeluPoTagu(context, TABLE_TAGS["04"]);
        if (t04) {
          const podaci = napravi0405Podatke(sveske, "04");
          await popuniTabelu(context, t04, podaci, 3);
          await sakrijTag(context, t04, TABLE_TAGS["04"]);
          ukupnoTabela++;
          console.log("✅ Tabela 04 generisana");
        }
      } catch(e) { console.error("❌ Tabela 04:", e.message); }

      // ---- TABELA 05 ----
      try {
        const t05 = await nadjiTabeluPoTagu(context, TABLE_TAGS["05"]);
        if (t05) {
          const podaci = napravi0405Podatke(sveske, "05");
          await popuniTabelu(context, t05, podaci, 3);
          await sakrijTag(context, t05, TABLE_TAGS["05"]);
          ukupnoTabela++;
          console.log("✅ Tabela 05 generisana");
        }
      } catch(e) { console.error("❌ Tabela 05:", e.message); }

      // ---- TABELA 061 ----
      try {
        const t061 = await nadjiTabeluPoTagu(context, TABLE_TAGS["061"]);
        if (t061) {
          const podaci = napravi061Podatke(projekti);
          await popuniTabelu(context, t061, podaci, 2);
          await sakrijTag(context, t061, TABLE_TAGS["061"]);
          ukupnoTabela++;
          console.log("✅ Tabela 061 generisana");
        }
      } catch(e) { console.error("❌ Tabela 061:", e.message); }

      // ---- TABELA 062 ----
      try {
        const t062 = await nadjiTabeluPoTagu(context, TABLE_TAGS["062"]);
        if (t062) {
          const podaci = napravi062Podatke(elaborati);
          await popuniTabelu(context, t062, podaci, 2);
          await sakrijTag(context, t062, TABLE_TAGS["062"]);
          ukupnoTabela++;
          console.log("✅ Tabela 062 generisana");
        }
      } catch(e) { console.error("❌ Tabela 062:", e.message); }

      // ---- TABELA 08 ----
      try {
        const t08 = await nadjiTabeluPoTagu(context, TABLE_TAGS["08"]);
        if (t08) {
          const podaci = napravi08Podatke(sveske);
          await popuniTabelu(context, t08, podaci, 1);
          await sakrijTag(context, t08, TABLE_TAGS["08"]);
          ukupnoTabela++;
          console.log("✅ Tabela 08 generisana");
        }
      } catch(e) { console.error("❌ Tabela 08:", e.message); }

      await context.sync();
    });

    if (ukupnoTabela === 0) {
      setStatus("⚠️ Nisu nađene placeholder tabele u dokumentu.", "warn");
    } else {
      setStatus(`✅ Generisano ${ukupnoTabela} tabele/a u dokumentu.`, "success");
    }

    closeTabeleModal();

  } catch (err) {
    console.error("❌ Greška pri generisanju tabela:", err);
    setStatus("Greška pri generisanju tabela. Vidi konzolu.", "error");
    if (btn) { btn.disabled = false; btn.textContent = "✅ Generiši tabele u Word-u"; }
  }
}

// ---- Podaci za tabelu 0.4 i 0.5 ----
// 3 kolone: oznaka | naziv | prazno (za licencu)
// ID tag ide u prvu ćeliju prvog reda (sakriven)
function napravi0405Podatke(sveske, tip) {
  return sveske.map((sv, i) => {
    const col0 = i === 0 ? `${TABLE_TAGS[tip]} ${sv.oznaka}` : sv.oznaka;
    const col1 = sv.tip === "glavna"
      ? "GLAVNA SVESKA"
      : sv.tip === "elaborat"
        ? sv.naziv
        : `PROJEKAT ${sv.naziv}`;
    return [col0, col1, ""];
  });
}

// ---- Podaci za tabelu 0.6.1 ----
// 2 kolone, 4 reda po projektu
function napravi061Podatke(projekti) {
  const rows = [];
  projekti.forEach((sv, pi) => {
    rows.push([
      pi === 0 ? `${TABLE_TAGS["061"]} ${sv.oznaka}. PROJEKAT ${sv.naziv}` : `${sv.oznaka}. PROJEKAT ${sv.naziv}`,
      ""
    ]);
    rows.push(["Projektant:", ""]);
    rows.push(["Broj licence:", ""]);
    rows.push(["Potpis:", ""]);
  });
  return rows;
}

// ---- Podaci za tabelu 0.6.2 ----
// 2 kolone, 5 redova po elaboratu
function napravi062Podatke(elaborati) {
  const rows = [];
  elaborati.forEach((sv, ei) => {
    rows.push([
      ei === 0 ? `${TABLE_TAGS["062"]} ${sv.naziv}:` : `${sv.naziv}:`,
      ""
    ]);
    rows.push(["Izrađivač:", ""]);
    rows.push(["Ovlašćeno lice:", ""]);
    rows.push(["Broj ovlašćenja:", ""]);
    rows.push(["Potpis:", ""]);
  });
  return rows;
}

// ---- Podaci za tabelu 0.8 ----
// 1 kolona, jedan red po grupi
function napravi08Podatke(sveske) {
  const seen = new Set();
  const rows = [];
  let rbr = 1;

  for (const sv of sveske) {
    if (sv.tip === "glavna") continue;
    const grupa = sv.oznaka.includes(".") ? sv.oznaka.split(".")[0] : sv.oznaka;
    if (seen.has(grupa)) continue;
    seen.add(grupa);

    const naslov = TABELE_08_NAZIVI[grupa] || sv.naziv;
    const tekst = `0.8.${rbr}. ${naslov}`;
    rows.push([rbr === 1 ? `${TABLE_TAGS["08"]} ${tekst}` : tekst]);
    rbr++;
  }
  return rows;
}

// ============================================
// HELPER: Upute za placeholder tabele
// ============================================
// Dodaje info u help text o tome kako da se postave placeholder tabele
// (poziva se jednom na inicijalizaciji)
function addTabeleHelpInfo() {
  const help = document.querySelector(".help");
  if (!help) return;
  const extra = document.createElement("div");
  extra.style.cssText = "margin-top:10px;border-top:1px solid #e5e7eb;padding-top:10px;";
  extra.innerHTML = `
    <strong>TABELE:</strong> U template ubaci običnu Word tabelu i u prvu ćeliju prvog reda upiši odgovarajući tag:<br>
    <code style="background:#f3f4f6;padding:1px 4px;border-radius:3px;font-size:11px;">[BA:04]</code> za tabelu 0.4,
    <code style="background:#f3f4f6;padding:1px 4px;border-radius:3px;font-size:11px;">[BA:05]</code> za 0.5,
    <code style="background:#f3f4f6;padding:1px 4px;border-radius:3px;font-size:11px;">[BA:061]</code>,
    <code style="background:#f3f4f6;padding:1px 4px;border-radius:3px;font-size:11px;">[BA:062]</code>,
    <code style="background:#f3f4f6;padding:1px 4px;border-radius:3px;font-size:11px;">[BA:08]</code>
    — tag se automatski sakriva pri generisanju.
  `;
  help.appendChild(extra);
}

// Pozovi help info na load
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", addTabeleHelpInfo);
} else {
  addTabeleHelpInfo();
}