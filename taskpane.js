/* global Office, Word */

// ============================================
// VERZIJA: KLIJENT - 2026-03-01
// Uprošćena verzija za klijente - samo POPUNI i OČISTI
// ============================================
console.log("🔧 BiroA Word Add-in KLIJENT VERZIJA: 2026-03-01");

let rows = [];

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
    const cleanValue = String(v).replace(/[^\d.,-]/g, "");
    const n = Number(cleanValue.replace(/\./g, "").replace(",", "."));
    if (Number.isNaN(n)) return v;

    const formatNumber = (num, decimals = 0) => {
      const fixed = num.toFixed(decimals);
      const parts = fixed.split(".");
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
      return decimals > 0 ? parts[0] + "," + parts[1] : parts[0];
    };

    if (format === "number:int") return formatNumber(Math.round(n), 0);
    if (format === "number:2") return formatNumber(n, 2);
    if (format === "number:rsd") return formatNumber(n, 2) + " RSD";
    if (format === "number:eur") return formatNumber(n, 2) + " €";
    if (format === "number:usd") return formatNumber(n, 2) + " $";
    if (format === "number:currency") return formatNumber(n, 2) + " RSD";
    return String(n);
  }

  if (type === "date") {
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
      let d;
      if (v.includes(".")) {
        const parts = v.split(".");
        if (parts.length === 3) d = new Date(parts[2], parts[1] - 1, parts[0]);
      } else if (v.includes("-")) {
        d = new Date(v);
      } else {
        return v;
      }
      if (!d || isNaN(d.getTime())) return v;
      const monthName = months[d.getMonth()];
      const yyyy = d.getFullYear();
      if (format === "date:mmmm.yyyy") return `${monthName}.${yyyy}`;
      const dd = String(d.getDate()).padStart(2, "0");
      return `${dd}.${monthName}.${yyyy}`;
    }

    return v;
  }

  if (type === "text") {
    if (format === "text:upper") return v.toUpperCase();
    if (format === "text:lower") return v.toLowerCase();
    if (format === "text:title") return v.replace(/\b\w/g, (l) => l.toUpperCase());
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

// ---------- render rows - klijent verzija (bez drag handle, bez delete, bez settings) ----------
function renderRows() {
  const container = el("rows");
  if (!container) return;

  container.innerHTML = "";

  if (rows.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        Nema polja u ovom dokumentu.<br>
        <span style="font-size:12px; color:#9ca3af;">Kontaktirajte projektanta za šablon.</span>
      </div>`;
    return;
  }

  rows.forEach((r) => {
    const row = document.createElement("div");
    row.className = "row client-row";

    // Field name (read-only)
    const fieldCell = document.createElement("div");
    fieldCell.className = "cell";
    const fieldLabel = document.createElement("div");
    fieldLabel.className = "field-label";
    fieldLabel.textContent = r.field || "";
    fieldCell.appendChild(fieldLabel);

    // Value input (editable)
    const valueCell = document.createElement("div");
    valueCell.className = "cell";
    const valueInput = document.createElement("input");
    valueInput.type = "text";
    valueInput.placeholder = r.type === "date" ? "dd.mm.yyyy" : r.type === "number" ? "0" : "Vrednost...";
    valueInput.value = r.value || "";
    valueInput.addEventListener("input", (e) => {
      r.value = e.target.value;
      saveValueToDocument();
    });
    valueCell.appendChild(valueInput);

    row.appendChild(fieldCell);
    row.appendChild(valueCell);
    container.appendChild(row);
  });
}

// ---------- XML state (čita/upisuje samo vrednosti) ----------
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

// Klijent samo upisuje vrednosti, ne može da menja strukturu
async function saveValueToDocument() {
  const xml = buildStateXml();
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

      context.document.customXmlParts.add(xml);
      await context.sync();
    });
  } catch (e) {
    console.warn("⚠️ Čuvanje vrednosti:", e.message);
  }
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
        id: crypto.randomUUID(),
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

// ---------- Word operations - samo popuni i očisti ----------
async function fillFieldsFromTable() {
  console.log("🔵 fillFieldsFromTable() KLIJENT");

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
      cc.insertText(out, Word.InsertLocation.replace);
      filled++;
    }

    await context.sync();
    console.log(`✅ Popunjeno ${filled} polja`);
    setStatus(`✅ Popunjeno ${filled} polja.`, "success");
  });

  await saveValueToDocument();
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
    setStatus(`Prikazana polja — {PLACEHOLDER} vraćen u ${cleared} mesta.`, "info");
  });
}

// ============================================
// GITHUB TEMPLATE PICKER
// ============================================

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

let _pickerSelectedBranch = null;
let _pickerSelectedFile = null;
let _strukeState = {};

function initStrukeState() {
  _strukeState = {};
  for (const s of STRUKE_LIST) {
    if (s.grupa === "00") continue;
    if (!_strukeState[s.grupa]) _strukeState[s.grupa] = [];
    _strukeState[s.grupa].push({ naziv: s.naziv, checked: false, custom: false });
  }
}

function buildGitHubRawUrl(branchId, fileName) {
  return `${GITHUB_CONFIG.baseUrl}/${branchId.split("/").map(encodeURIComponent).join("/")}/${encodeURIComponent(fileName)}`;
}

async function downloadFileContent(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`GitHub download greška: ${response.status}`);
  return await response.arrayBuffer();
}

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
    btn.addEventListener("mouseenter", () => { btn.style.background = "#eff6ff"; btn.style.borderColor = "#93c5fd"; });
    btn.addEventListener("mouseleave", () => { btn.style.background = "#f9fafb"; btn.style.borderColor = "#e5e7eb"; });
    btn.addEventListener("click", () => {
      _pickerSelectedBranch = branch;
      renderPickerStep2();
    });
    body.appendChild(btn);
  });

  setPickerFooter([{ label: "Otkaži", onClick: closeGitHubTemplateModal }]);
}

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
    btn.addEventListener("mouseenter", () => { btn.style.background = "#eff6ff"; btn.style.borderColor = "#93c5fd"; });
    btn.addEventListener("mouseleave", () => { btn.style.background = "#f9fafb"; btn.style.borderColor = "#e5e7eb"; });
    btn.addEventListener("click", () => {
      _pickerSelectedFile = file;
      openTemplateFromGitHub();
    });
    body.appendChild(btn);
  });

  setPickerFooter([{ label: "Otkaži", onClick: closeGitHubTemplateModal }]);
}

function extractGrupaNaslov(stavke) {
  if (stavke.length === 1) return stavke[0].naziv;
  const first = stavke[0].naziv;
  const words = first.split(" ");
  let common = "";
  for (let i = words.length; i > 0; i--) {
    const candidate = words.slice(0, i).join(" ");
    if (stavke.every(s => s.naziv.startsWith(candidate))) { common = candidate; break; }
  }
  return common || first;
}

function createStrukaRow(grupa, idx, stavka) {
  const row = document.createElement("div");
  row.style.cssText = `display:flex; align-items:center; gap:8px; padding:6px 12px; border-top:1px solid #f3f4f6;`;
  row.dataset.idx = idx;

  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.checked = stavka.checked;
  cb.style.cssText = "width:15px;height:15px;cursor:pointer;accent-color:#1d4ed8;";
  cb.addEventListener("change", () => { _strukeState[grupa][idx].checked = cb.checked; });

  const label = document.createElement("span");
  label.style.cssText = "flex:1; font-size:12px; color:#374151;";
  label.textContent = stavka.naziv;

  if (stavka.custom) {
    const delBtn = document.createElement("button");
    delBtn.textContent = "×";
    delBtn.style.cssText = `background:none; border:none; color:#9ca3af; font-size:16px; cursor:pointer; padding:0 4px; line-height:1;`;
    delBtn.addEventListener("click", () => {
      _strukeState[grupa].splice(idx, 1);
      const groupBody = document.getElementById(`strukaGroup_${grupa}`);
      if (groupBody) {
        groupBody.innerHTML = "";
        _strukeState[grupa].forEach((s, i) => groupBody.appendChild(createStrukaRow(grupa, i, s)));
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

function addCustomStavka(grupa, groupBody) {
  const formRow = document.createElement("div");
  formRow.style.cssText = "display:flex;gap:6px;padding:6px 12px;border-top:1px solid #f3f4f6;";

  const input = document.createElement("input");
  input.type = "text";
  input.placeholder = "Naziv sveske...";
  input.style.cssText = `flex:1; padding:4px 8px; border:1px solid #93c5fd; border-radius:5px; font-size:12px; outline:none;`;

  const confirmBtn = document.createElement("button");
  confirmBtn.textContent = "Dodaj";
  confirmBtn.style.cssText = `padding:4px 10px; background:#1d4ed8; color:#fff; border:none; border-radius:5px; font-size:12px; cursor:pointer;`;

  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "×";
  cancelBtn.style.cssText = `padding:4px 8px; background:#f3f4f6; border:none; border-radius:5px; font-size:12px; cursor:pointer;`;

  const doAdd = () => {
    const naziv = input.value.trim();
    if (!naziv) { input.focus(); return; }
    _strukeState[grupa].push({ naziv, checked: true, custom: true });
    groupBody.innerHTML = "";
    _strukeState[grupa].forEach((s, i) => groupBody.appendChild(createStrukaRow(grupa, i, s)));
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
    setStatus(`✅ Otvoren: ${fileName}`, "success");

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

// ---------- Tab switching ----------
function setupTabs() {
  const btnFill = el("btnFill");
  const btnClear = el("btnClear");
  const btnTemplates = el("btnTemplates");

  if (btnFill) {
    btnFill.addEventListener("click", async () => {
      btnFill.disabled = true;
      btnFill.textContent = "⏳ Popunjavam...";
      try {
        await fillFieldsFromTable();
      } catch (e) {
        console.error("❌ Greška pri popunjavanju:", e);
        setStatus("Greška pri popunjavanju. Vidi konzolu.", "error");
      } finally {
        btnFill.disabled = false;
        btnFill.textContent = "POPUNI";
      }
    });
  }

  if (btnClear) {
    btnClear.addEventListener("click", async () => {
      btnClear.disabled = true;
      try {
        await clearFieldsKeepControls();
      } catch (e) {
        console.error("❌ Greška pri čišćenju:", e);
        setStatus("Greška pri čišćenju.", "error");
      } finally {
        btnClear.disabled = false;
      }
    });
  }

  if (btnTemplates) {
    btnTemplates.addEventListener("click", () => {
      openGitHubTemplateModal();
    });
  }
}

// ---------- Init ----------
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) {
    document.body.innerHTML = `
      <div style="padding:20px; color:#dc2626; font-family: 'Segoe UI', sans-serif;">
        Ovaj dodatak radi samo u Microsoft Word-u.
      </div>`;
    return;
  }

  console.log("✅ Office je spreman - klijent verzija");

  // Učitaj stanje iz dokumenta
  try {
    await loadStateFromDocument();
    renderRows();

    if (rows.length === 0) {
      setStatus("Dokument nema sačuvana polja. Otvorite šablon koji je projektant pripremio.", "warn");
    } else {
      setStatus(`Učitano ${rows.length} polja iz dokumenta.`, "success");
    }
  } catch (e) {
    console.error("❌ Greška pri učitavanju:", e);
    setStatus("Greška pri učitavanju dokumenta.", "error");
    renderRows();
  }

  setupTabs();
});
