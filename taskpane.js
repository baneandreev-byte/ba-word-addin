/* global Office, Word */

// ============================================
// VERZIJA: 2025-02-07 - V34 (STABILNI DRAG & DROP)
// ============================================
console.log("🔧 BA Word Add-in VERZIJA: 2025-02-07 - V34");
console.log("✅ FIX: Drag & Drop koristi stabilne ID-ove umesto index-a");
console.log("✅ FIX: stopPropagation na ⚙ i × sprečava duhove klikova");
console.log("✅ FIX: Modal backdrop guard - klik samo na prazno zatvara modal");
console.log("✅ Sve stabilno i radi kako treba!");

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
  draggedElement = this;
  draggedId = this.dataset.id;
  
  this.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/html', this.innerHTML);
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
  
  // Update selected index if needed
  if (selectedRowIndex === fromIndex) {
    selectedRowIndex = toIndex;
  } else if (fromIndex < selectedRowIndex && toIndex >= selectedRowIndex) {
    selectedRowIndex--;
  } else if (fromIndex > selectedRowIndex && toIndex <= selectedRowIndex) {
    selectedRowIndex++;
  }
  
  // Re-render and save
  renderRows();
  saveStateToDocument();
  
  // Show status
  setStatus(`Polje "${movedItem.field}" premešteno.`, "info");
  
  return false;
}

function handleDragEnd(e) {
  this.classList.remove('dragging');
  
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
    
    // Make row draggable - koristi ID umesto index
    row.draggable = true;
    row.dataset.id = r.id;
    row.dataset.index = idx; // Zadrži index za backward compatibility
    
    // Drag event listeners
    row.addEventListener('dragstart', handleDragStart);
    row.addEventListener('dragover', handleDragOver);
    row.addEventListener('dragleave', handleDragLeave);
    row.addEventListener('drop', handleDrop);
    row.addEventListener('dragend', handleDragEnd);

    // Click handler na ceo red - selektuje red za ubacivanje
    row.addEventListener("click", (e) => {
      // Don't select if clicking drag handle
      if (e.target.closest('.drag-handle')) return;
      selectedRowIndex = idx;
      renderRows();
    });

    // Drag handle
    const dragHandle = document.createElement("div");
    dragHandle.className = "drag-handle";
    dragHandle.innerHTML = "⋮⋮";
    dragHandle.title = "Prevuci za premeštanje";

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
// FIX: deleteControlsAndXml - Custom Modal umesto confirm()
// ============================================
async function deleteControlsAndXml() {
  // Prikaži custom confirm modal
  showDeleteConfirmModal();
}

async function performDelete() {
  try {
    console.log("🔴 Počinjem brisanje content controls...");
    
    const map = buildValueMap();
    let removed = 0;

    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items");
      await context.sync();

      console.log(`  Pronađeno ${ccs.items.length} content controls`);

      // Učitaj sve potrebne properties
      for (const cc of ccs.items) {
        cc.load("tag,text");
      }
      await context.sync();

      // Prvo prolaz: zameni svaki CC sa plain text-om
      const toDelete = [];
      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta) continue;

        console.log(`  - Obrađujem CC za polje: ${meta.key}`);

        // Uzmi formatirani tekst ili trenutni tekst u CC-u
        const finalText = map.get(meta.key)?.formatted ?? cc.text;
        
        // Umetni plain text NAKON content control-a
        const range = cc.getRange(Word.RangeLocation.after);
        range.insertText(finalText, Word.InsertLocation.start);
        
        // Označi CC za brisanje
        toDelete.push(cc);
        removed++;
      }

      await context.sync();
      console.log(`  Umetnut tekst za ${removed} polja`);

      // Drugi prolaz: obriši sve CC-ove BEZ sadržaja (text je već van CC-a)
      for (const cc of toDelete) {
        try {
          cc.delete(false); // false jer smo vec izvukli text
        } catch (e) {
          console.error("  Greška pri brisanju:", e);
        }
      }

      await context.sync();
    });

    console.log(`✅ Uklonjeno ${removed} content controls`);

    // Obriši XML state
    await deleteSavedStateFromDocument();
    console.log("✅ XML state obrisan");

    // Očisti lokalne podatke
    rows = [];
    selectedRowIndex = null;
    renderRows();

    setStatus(`Dokument očišćen: ${removed} polja uklonjeno, plugin podaci obrisani.`, "info");
  } catch (error) {
    console.error("❌ GREŠKA pri brisanju:", error);
    setStatus("Greška pri brisanju polja. Vidi konzolu.", "error");
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

  // Delete modal buttons
  const btnDeleteModalClose = el("btnDeleteModalClose");
  const btnDeleteCancel = el("btnDeleteCancel");
  const btnDeleteConfirm = el("btnDeleteConfirm");

  if (btnInsert) btnInsert.addEventListener("click", insertFieldAtSelection);
  if (btnFill) btnFill.addEventListener("click", fillFieldsFromTable);
  if (btnClear) btnClear.addEventListener("click", clearFieldsKeepControls);
  if (btnDelete) btnDelete.addEventListener("click", deleteControlsAndXml);
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
  
  // Spreci da klik NA modal zatvara modal
  const modal = el("modal");
  const deleteModal = el("deleteModal");
  if (modal) modal.addEventListener("click", (e) => e.stopPropagation());
  if (deleteModal) deleteModal.addEventListener("click", (e) => e.stopPropagation());
  
  // Backdrop zatvara modal SAMO ako klikneš na backdrop (ne na modal)
  if (modalBackdrop) {
    modalBackdrop.addEventListener("click", (e) => {
      if (e.target !== modalBackdrop) return; // samo klik na "prazno"
      closeModal();
      closeDeleteModal();
    });
  }
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
