/* global Office, Word */

// ============================================
// VERZIJA: KLIJENT - 2026-03-01
// Klijent verzija: bez UBACI POLJE, bez TABELE, bez dodavanja/brisanja redova
// ============================================
console.log("🔧 BA Word Add-in KLIJENT VERZIJA: 2026-03-01");

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
    { value: "date:mmmm.yyyy", label: "MMMM yyyy", hint: "Primer: februar 2025" },
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
        return `${monthName} ${yyyy}`;
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
    container.innerHTML = '<div class="empty-state"><div class="empty-icon">📄</div><div>Nema učitanih polja.</div></div>';
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
    btnEdit.className = "btn-settings";
    btnEdit.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedRowIndex = idx;
      openModal(r);
    });

    actionsCell.appendChild(btnEdit);

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
    setStatus(`Prikazano ${cleared} polja — {PLACEHOLDER} vraćen.`, "info");
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
  const btnFill = el("btnFill");
  const btnClear = el("btnClear");
  const btnDelete = el("btnDelete");
  const btnTemplates = el("btnTemplates");
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

  // Active tab highlighting
  const tabBtns = [btnFill, btnClear, btnDelete, btnTemplates];
  function setActiveTab(active) {
    tabBtns.forEach(b => { if (b) b.classList.remove("active"); });
    if (active) active.classList.add("active");
  }

  if (btnFill) btnFill.addEventListener("click", () => { setActiveTab(btnFill); fillFieldsFromTable(); });
  if (btnClear) btnClear.addEventListener("click", () => { setActiveTab(btnClear); clearFieldsKeepControls(); });
  if (btnDelete) btnDelete.addEventListener("click", () => { deleteControlsAndXml(); });
  if (btnTemplates) btnTemplates.addEventListener("click", () => { setActiveTab(btnTemplates); openGitHubTemplateModal(); });
  if (btnExportCSV) btnExportCSV.addEventListener("click", exportCSV);
  if (btnImportCSV) btnImportCSV.addEventListener("click", importCSV);

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
      if (e.target !== modalBackdrop) return;
      closeModal();
      closeDeleteModal();
    });
  }
}

Office.onReady(async () => {
  console.log("✅ Office.onReady KLIJENT STARTED");
  
  try {
    await loadStateFromDocument();
    console.log("✅ loadStateFromDocument završen, rows.length:", rows.length);
  } catch (e) {
    console.error("❌ Load state error:", e);
  }

  renderRows();
  bindUi();
  
  console.log("✅✅✅ Office.onReady COMPLETED ✅✅✅");
});

// ============================================
