/* global Office, Word */

// ============================================
// VERZIJA: 2025-02-08 - V45 (IMPROVED DELETE)
// ============================================
console.log("🔧 BA Word Add-in VERZIJA: 2025-02-08 - V45");
console.log("✅ NOVO: Poboljšana delete funkcija - jednostavnija logika");
console.log("✅ SharePoint templejti - Graph API integracija");

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
    { value: "date:today", label: "Danas (dd.mm.yyyy)", hint: "Primer: 08.02.2025" },
    { value: "date:dd.mm.yyyy", label: "dd.mm.yyyy", hint: "Primer: 08.02.2025" },
    { value: "date:yyyy-mm-dd", label: "yyyy-mm-dd", hint: "Primer: 2025-02-08" },
    { value: "date:mmmm.yyyy", label: "MMMM.yyyy", hint: "Primer: februar.2025" },
    { value: "date:dd.mmmm.yyyy", label: "dd.MMMM.yyyy", hint: "Primer: 08.februar.2025" },
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
  const backdrop = el("modalBackdrop");
  if (!modal || !backdrop) return;
  
  modal.classList.remove("hidden");
  backdrop.classList.remove("hidden");
}

function closeDeleteModal() {
  const modal = el("deleteModal");
  const backdrop = el("modalBackdrop");
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
// ⭐ V45 - IMPROVED DELETE FUNCTION
// ============================================
async function deleteControlsAndXml() {
  // Prikaži custom confirm modal
  showDeleteConfirmModal();
}

/**
 * ⭐ POBOLJŠANA VERZIJA - Jednostavnija i bezbednija
 * Koristi cc.delete(false) direktno umesto kompleksnog insertText pristupa
 */
async function performDelete() {
  try {
    console.log("🔴 Počinjem brisanje content controls...");
    
    let removed = 0;

    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items");
      await context.sync();

      console.log(`📊 Pronađeno ${ccs.items.length} content controls`);

      if (ccs.items.length === 0) {
        console.log("ℹ️ Nema content control-a za brisanje");
        setStatus("Nema polja za brisanje.", "info");
        closeDeleteModal();
        return;
      }

      // FAZA 1: Učitaj properties za sve kontrole
      for (const cc of ccs.items) {
        cc.load("tag,text,cannotDelete");
      }
      await context.sync();
      console.log("✅ Properties učitane");

      // FAZA 2: Obriši BA_FIELD kontrole - iteracija unazad
      for (let i = ccs.items.length - 1; i >= 0; i--) {
        const cc = ccs.items[i];
        const meta = parseTag(cc.tag || "");
        
        // Preskači ako nije BA_FIELD
        if (!meta) {
          console.log(`  ⏭️ [${i}] Preskačem: nije BA_FIELD`);
          continue;
        }

        console.log(`  🔍 [${i}] Procesiranje: ${meta.key}`);

        // Otključaj ako je zaključana
        if (cc.cannotDelete) {
          console.log(`    🔓 Otključavanje kontrole`);
          cc.cannotDelete = false;
        }

        const currentText = cc.text || "";
        console.log(`    📝 Tekst: "${currentText}"`);

        // ⭐ KRITIČNA LINIJA: delete(false) = zadrži tekst
        // Ova metoda automatski zadržava formatiranje i poziciju teksta
        cc.delete(false);
        removed++;
        
        console.log(`    ✅ Kontrola obrisana (tekst zadržan)`);
      }

      await context.sync();
      console.log(`✅ Obrisano ${removed} kontrola`);
    });

    if (removed === 0) {
      setStatus("Nema BiroA polja za brisanje.", "info");
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

    setStatus(`Dokument očišćen: ${removed} polja uklonjeno, plugin podaci obrisani.`, "info");
    closeDeleteModal();
    
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
// TEMPLATE MANAGER (V30 - SharePoint Integration)
// ============================================

// SharePoint site configuration
const SHAREPOINT_CONFIG = {
  siteUrl: "https://biroa.sharepoint.com/sites/Officetamplates",
  folderPath: "/sites/Officetamplates/Deljeni dokumenti/Table addin word templetes"
};

let templates = [];
let editingTemplateId = null;

// ---------- Graph API Helpers ----------

// Get access token for Graph API
async function getGraphToken() {
  try {
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      forMSGraphAccess: true
    });
    return token;
  } catch (error) {
    console.error("❌ Greška pri dobijanju tokena:", error);
    console.error("❌ Error code:", error.code);
    console.error("❌ Error message:", error.message);
    console.error("❌ Error name:", error.name);
    
    // Detaljne poruke za različite greške
    if (error.code === 13001) {
      throw new Error("CONSENT REQUIRED: Korisnik mora da odobri pristup. Klikni 'Allow' kada se pojavi popup.");
    } else if (error.code === 13002) {
      throw new Error("USER NOT SIGNED IN: Korisnik nije ulogovan u Office. Uloguj se u Word.");
    } else if (error.code === 13003) {
      throw new Error("INTERNAL ERROR: Office SSO greška. Restartuj Word i probaj ponovo.");
    } else if (error.code === 13004) {
      throw new Error("INVALID RESOURCE: Resource URL u manifestu je pogrešan.");
    } else if (error.code === 13005) {
      throw new Error("INVALID GRANT: Token je istekao ili je nevažeći.");
    } else if (error.code === 13006) {
      throw new Error("CLIENT ERROR: Greška u konfiguraciji Azure AD aplikacije.");
    } else if (error.code === 13007) {
      throw new Error("MISSING CONSENT: Admin consent nije dat za aplikaciju.");
    } else if (error.code === 13012) {
      throw new Error("POPUP BLOCKED: Consent popup je blokiran. Dozvoli popups.");
    } else {
      throw new Error(`SSO greška (${error.code || 'unknown'}): ${error.message || 'Proveri Azure AD konfiguraciju'}`);
    }
  }
}

// Call Graph API
async function callGraphAPI(endpoint, method = "GET", body = null) {
  try {
    const token = await getGraphToken();
    
    const options = {
      method: method,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };
    
    if (body) {
      options.body = JSON.stringify(body);
    }
    
    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, options);
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error("Graph API greška:", response.status, errorText);
      throw new Error(`Graph API greška: ${response.status}`);
    }
    
    return await response.json();
  } catch (error) {
    console.error("❌ callGraphAPI greška:", error);
    throw error;
  }
}

// Get SharePoint site ID
async function getSharePointSiteId() {
  try {
    // Extract hostname and site path from URL
    const url = new URL(SHAREPOINT_CONFIG.siteUrl);
    const hostname = url.hostname;
    const sitePath = url.pathname;
    
    // Call Graph API to get site
    const site = await callGraphAPI(`/sites/${hostname}:${sitePath}`);
    return site.id;
  } catch (error) {
    console.error("❌ getSharePointSiteId greška:", error);
    throw error;
  }
}

// Get files from SharePoint folder
async function getSharePointFiles(folderPath) {
  try {
    const siteId = await getSharePointSiteId();
    
    // Encode folder path
    const encodedPath = encodeURIComponent(folderPath);
    
    // Get drive items
    const result = await callGraphAPI(`/sites/${siteId}/drive/root:${encodedPath}:/children`);
    
    // Filter only .docx files
    const docxFiles = result.value.filter(file => 
      file.name.toLowerCase().endsWith('.docx') && !file.name.startsWith('~')
    );
    
    console.log(`✅ Pronađeno ${docxFiles.length} .docx fajlova`);
    return docxFiles;
  } catch (error) {
    console.error("❌ getSharePointFiles greška:", error);
    throw error;
  }
}

// Download file content from SharePoint
async function downloadFileContent(fileId) {
  try {
    const siteId = await getSharePointSiteId();
    const token = await getGraphToken();
    
    // Get download URL
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/content`,
      {
        headers: { 'Authorization': `Bearer ${token}` }
      }
    );
    
    if (!response.ok) {
      throw new Error(`Download greška: ${response.status}`);
    }
    
    return await response.arrayBuffer();
  } catch (error) {
    console.error("❌ downloadFileContent greška:", error);
    throw error;
  }
}

// Extract fields from .docx file
async function extractFieldsFromDocx(arrayBuffer) {
  try {
    // Load the docx file using JSZip
    const zip = await JSZip.loadAsync(arrayBuffer);
    
    // Read document.xml
    const docXml = await zip.file("word/document.xml").async("string");
    
    // Parse XML
    const parser = new DOMParser();
    const doc = parser.parseFromString(docXml, "text/xml");
    
    // Find all content controls with BA_FIELD tags
    const controls = doc.querySelectorAll('w\\:tag, tag');
    const fields = [];
    
    controls.forEach(tagNode => {
      const tagValue = tagNode.getAttribute('w:val') || tagNode.textContent;
      const parsed = parseTag(tagValue);
      
      if (parsed) {
        // Check if field already exists
        if (!fields.find(f => f.field === parsed.key)) {
          fields.push({
            field: parsed.key,
            type: parsed.type,
            format: parsed.format
          });
        }
      }
    });
    
    console.log(`✅ Ekstraktovano ${fields.length} polja iz dokumenta`);
    return fields;
  } catch (error) {
    console.error("❌ extractFieldsFromDocx greška:", error);
    return [];
  }
}

// ---------- Template Management ----------

// Učitaj templejte sa SharePointa
async function loadTemplatesFromSharePoint() {
  try {
    setStatus("Učitavam templejte sa SharePointa...", "info");
    
    const files = await getSharePointFiles(SHAREPOINT_CONFIG.folderPath);
    
    templates = files.map(file => ({
      id: file.id,
      name: file.name.replace('.docx', ''),
      desc: `SharePoint: ${new Date(file.lastModifiedDateTime).toLocaleDateString('sr-RS')}`,
      fileId: file.id,
      downloadUrl: file['@microsoft.graph.downloadUrl'],
      fields: [] // Will be loaded on demand
    }));
    
    console.log("✅ Učitano", templates.length, "templata sa SharePointa");
    setStatus(`Učitano ${templates.length} templata`, "success");
  } catch (error) {
    console.error("❌ Greška pri učitavanju templata:", error);
    setStatus("Greška pri učitavanju templata sa SharePointa", "error");
    templates = [];
    
    // Fallback to local XML if SharePoint fails
    console.log("⚠️ Pokušavam da učitam lokalne templejte...");
    await loadTemplatesFromDocument();
  }
}

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
            const fieldNodes = fieldsNode.querySelectorAll("field");
            fieldNodes.forEach((fn) => {
              fields.push({
                field: fn.getAttribute("name") || "",
                type: fn.getAttribute("type") || "text",
                format: fn.getAttribute("format") || "text:auto",
              });
            });
          }
          
          templates.push({ id, name, desc, fields });
        });
        
        console.log("✅ Učitano", templates.length, "lokalnih templata");
      }
    });
  } catch (error) {
    console.error("❌ Greška pri učitavanju lokalnih templata:", error);
  }
}

// Sačuvaj templejte u dokument (fallback za lokalno čuvanje)
async function saveTemplatesToDocument() {
  try {
    await Word.run(async (context) => {
      const targetNamespace = "http://biroa.com/word-addin/templates";
      const parts = context.document.customXmlParts;
      parts.load("items");
      await context.sync();

      // Obriši stare templejte
      for (const p of parts.items) {
        p.load("namespaceUri");
      }
      await context.sync();
      
      for (const p of parts.items) {
        if (p.namespaceUri === targetNamespace) p.delete();
      }
      await context.sync();

      // Kreiraj novi XML
      let xml = `<templates xmlns="${targetNamespace}">`;
      templates.forEach((t) => {
        const nameEsc = xmlEscape(t.name);
        const descEsc = xmlEscape(t.desc || "");
        xml += `<template id="${t.id}" name="${nameEsc}" description="${descEsc}">`;
        xml += `<fields>`;
        (t.fields || []).forEach((f) => {
          const fNameEsc = xmlEscape(f.field);
          const fTypeEsc = xmlEscape(f.type || "text");
          const fFormatEsc = xmlEscape(f.format || "text:auto");
          xml += `<field name="${fNameEsc}" type="${fTypeEsc}" format="${fFormatEsc}"/>`;
        });
        xml += `</fields>`;
        xml += `</template>`;
      });
      xml += `</templates>`;

      context.document.customXmlParts.add(xml);
      await context.sync();
      
      console.log("✅ Templejti sačuvani lokalno");
    });
  } catch (error) {
    console.error("❌ Greška pri čuvanju templata:", error);
  }
}

// Primeni templejt - postavi rows iz templata
async function applyTemplate(templateId) {
  const template = templates.find((t) => t.id === templateId);
  if (!template) return;
  
  try {
    setStatus(`Učitavam templejt "${template.name}"...`, "info");
    
    // If it's a SharePoint template and fields are not loaded yet
    if (template.fileId && (!template.fields || template.fields.length === 0)) {
      const arrayBuffer = await downloadFileContent(template.fileId);
      template.fields = await extractFieldsFromDocx(arrayBuffer);
    }
    
    if (!template.fields || template.fields.length === 0) {
      setStatus("Templejt nema polja", "error");
      return;
    }
    
    // Set rows
    rows = template.fields.map((f) => ({
      id: crypto.randomUUID(),
      field: f.field,
      value: "",
      type: f.type || "text",
      format: f.format || "text:auto",
    }));
    
    selectedRowIndex = null;
    renderRows();
    await saveStateToDocument();
    
    setStatus(`Učitan templejt "${template.name}" (${rows.length} polja)`, "success");
    closeTemplatesModal();
  } catch (error) {
    console.error("❌ Greška pri primeni templata:", error);
    setStatus("Greška pri učitavanju templata", "error");
  }
}

// Obriši templejt
async function deleteTemplate(templateId) {
  const template = templates.find((t) => t.id === templateId);
  if (!template) return;
  
  if (!confirm(`Obrisati templejt "${template.name}"?`)) return;
  
  templates = templates.filter((t) => t.id !== templateId);
  await saveTemplatesToDocument();
  setStatus(`Templejt "${template.name}" obrisan`, "success");
  renderTemplatesList();
}

// Otvori modal sa templatejima
function openTemplatesModal() {
  const backdrop = el("modalTemplatesBackdrop");
  const modal = el("modalTemplates");
  if (!backdrop || !modal) return;
  
  backdrop.classList.remove("hidden");
  modal.classList.remove("hidden");
  
  renderTemplatesList();
}

// Zatvori modal sa templatejima
function closeTemplatesModal() {
  const backdrop = el("modalTemplatesBackdrop");
  const modal = el("modalTemplates");
  if (backdrop) backdrop.classList.add("hidden");
  if (modal) modal.classList.add("hidden");
}

// Renderuj listu templata
function renderTemplatesList() {
  const container = el("templatesList");
  const searchInput = el("templateSearch");
  if (!container) return;
  
  const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : "";
  
  // Filter templates
  const filtered = templates.filter((t) => {
    if (!searchTerm) return true;
    return (
      t.name.toLowerCase().includes(searchTerm) ||
      (t.desc && t.desc.toLowerCase().includes(searchTerm))
    );
  });
  
  container.innerHTML = "";
  
  if (filtered.length === 0) {
    container.innerHTML = '<div class="empty-state">Nema templata</div>';
    return;
  }
  
  filtered.forEach((template) => {
    const card = document.createElement("div");
    card.className = "template-card";
    
    // Header
    const header = document.createElement("div");
    header.className = "template-card-header";
    
    const title = document.createElement("h4");
    title.className = "template-card-title";
    title.textContent = template.name;
    
    const actions = document.createElement("div");
    actions.className = "template-card-actions";
    
    // Edit button (only for local templates)
    if (!template.fileId) {
      const btnEdit = document.createElement("button");
      btnEdit.className = "template-card-btn";
      btnEdit.innerHTML = "✎";
      btnEdit.title = "Izmeni";
      btnEdit.addEventListener("click", (e) => {
        e.stopPropagation();
        openEditTemplateModal(template.id);
      });
      actions.appendChild(btnEdit);
    }
    
    // Delete button (only for local templates)
    if (!template.fileId) {
      const btnDelete = document.createElement("button");
      btnDelete.className = "template-card-btn delete";
      btnDelete.innerHTML = "🗑️";
      btnDelete.title = "Obriši";
      btnDelete.addEventListener("click", (e) => {
        e.stopPropagation();
        deleteTemplate(template.id);
      });
      actions.appendChild(btnDelete);
    }
    
    header.appendChild(title);
    header.appendChild(actions);
    
    // Description
    const desc = document.createElement("div");
    desc.className = "template-card-desc";
    desc.textContent = template.desc || "(bez opisa)";
    
    // Meta
    const meta = document.createElement("div");
    meta.className = "template-card-meta";
    
    const fieldsCount = document.createElement("span");
    fieldsCount.className = "template-card-fields";
    fieldsCount.textContent = template.fields ? 
      `${template.fields.length} polja` : 
      (template.fileId ? "SharePoint dokument" : "0 polja");
    
    meta.appendChild(fieldsCount);
    
    // Click handler
    card.addEventListener("click", () => {
      applyTemplate(template.id);
    });
    
    card.appendChild(header);
    card.appendChild(desc);
    card.appendChild(meta);
    
    container.appendChild(card);
  });
}

// Otvori modal za editovanje templata
function openEditTemplateModal(templateId) {
  const backdrop = el("modalEditTemplateBackdrop");
  const modal = el("modalEditTemplate");
  const title = el("editTemplateTitle");
  const nameInput = el("templateName");
  const descInput = el("templateDesc");
  const fieldsList = el("templateFieldsList");
  
  if (!backdrop || !modal) return;
  
  editingTemplateId = templateId;
  
  if (templateId) {
    // Editing existing template
    const template = templates.find((t) => t.id === templateId);
    if (!template) return;
    
    if (title) title.textContent = "Izmeni templejt";
    if (nameInput) nameInput.value = template.name;
    if (descInput) descInput.value = template.desc || "";
    
    // Show template fields
    if (fieldsList) {
      fieldsList.innerHTML = "";
      if (!template.fields || template.fields.length === 0) {
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
  if (btnTemplates) btnTemplates.addEventListener("click", openTemplatesModal);
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
  try {
    await loadStateFromDocument();
    await loadTemplatesFromSharePoint(); // Load from SharePoint
  } catch (e) {
    console.error("Load state error:", e);
  }

  renderRows();
  bindUi();
});
