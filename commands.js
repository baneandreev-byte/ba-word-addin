/* global Office, Word */

// ============================================
// VERZIJA: 2025-02-10 - V47 (WITH CONFIRMATION)
// commands.js - Ribbon Command Functions
// ============================================
console.log("🔧 BA Word Add-in Commands VERZIJA: 2025-02-10 - V47");
console.log("✅ SA CONFIRMATION DIALOG-OM");
console.log("✅ Detaljno mapiranje pre brisanja");

/**
 * Parse BA_FIELD tag to extract metadata
 * Format: BA_FIELD|key=NAZIV|type=text|format=text:auto
 */
function parseTag(tag) {
  const s = String(tag || "");
  
  if (!s.startsWith("BA_FIELD|")) {
    return null;
  }
  
  const parts = s.split("|").slice(1);
  const out = {};
  
  for (const p of parts) {
    const [k, ...rest] = p.split("=");
    out[k] = rest.join("=");
  }
  
  if (!out.key) {
    return null;
  }
  
  return {
    key: out.key,
    type: out.type || "text",
    format: out.format || "text:auto",
  };
}

/**
 * Delete XML custom parts that store plugin state
 */
async function deleteXMLState(context) {
  try {
    const parts = context.document.customXmlParts;
    parts.load("items");
    await context.sync();

    const toDelete = [];
    for (const part of parts.items) {
      part.load("namespaceUri");
    }
    await context.sync();

    for (const part of parts.items) {
      if (part.namespaceUri === "http://biroa.rs/word-addin/state") {
        toDelete.push(part);
      }
    }

    for (const part of toDelete) {
      part.delete();
    }
    
    if (toDelete.length > 0) {
      await context.sync();
      console.log(`✅ Obrisano ${toDelete.length} XML custom parts`);
    }
  } catch (error) {
    console.error("⚠️ Greška pri brisanju XML state:", error);
  }
}

/**
 * 📋 FAZA 1: Mapiranje svih kontrola u dokumentu
 * Analizira kontrole i vraća podatke za confirmation dialog
 */
async function mapContentControls() {
  console.log("🔄 FAZA 1: Mapiranje content controls...");
  console.log("=".repeat(60));
  
  const mappedControls = [];
  let totalControls = 0;
  let skippedControls = 0;

  await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    totalControls = contentControls.items.length;
    console.log(`📊 Pronađeno ${totalControls} content controls u dokumentu`);

    if (totalControls === 0) {
      console.log("ℹ️ Nema content control-a");
      return;
    }

    // Učitaj properties za sve kontrole
    for (const cc of contentControls.items) {
      cc.load("tag,text,title");
    }
    await context.sync();
    console.log("✅ Properties učitane");

    // Analiziraj sve kontrole
    console.log("\n📋 Detaljno mapiranje:\n" + "-".repeat(60));
    
    for (let i = 0; i < contentControls.items.length; i++) {
      const cc = contentControls.items[i];
      const tag = cc.tag || "";
      const title = cc.title || "(bez naslova)";
      const text = cc.text || "";
      
      console.log(`\n[${i}] Kontrola:`);
      console.log(`    Title: "${title}"`);
      console.log(`    Tag: "${tag}"`);
      console.log(`    Text: "${text.substring(0, 80)}${text.length > 80 ? '...' : ''}"`);
      
      const meta = parseTag(tag);
      
      if (!meta) {
        console.log(`    ⏭️ PRESKAČEM - nije BA_FIELD`);
        skippedControls++;
        continue;
      }
      
      console.log(`    ✅ MAPIRAN - BA_FIELD kontrola`);
      console.log(`    📝 Tekst koji će biti zadržan: "${text}"`);
      
      // Dodaj u listu za brisanje
      mappedControls.push({
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
    console.log(`   BA_FIELD kontrola: ${mappedControls.length}`);
    console.log(`   Preskočeno: ${skippedControls}`);
    console.log("=".repeat(60));
  });

  return {
    controls: mappedControls,
    total: totalControls,
    skipped: skippedControls
  };
}

/**
 * 🗑️ FAZA 2: Brisanje kontrola nakon potvrde
 * Prima listu kontrola iz mapiranja i briše ih
 */
async function deleteControlsByIndices(controlIndices) {
  console.log("\n🔄 FAZA 2: Brisanje potvđenih kontrola...");
  console.log("=".repeat(60));
  
  let removed = 0;

  await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    // Učitaj properties
    for (const cc of contentControls.items) {
      cc.load("tag,text,cannotDelete");
    }
    await context.sync();

    console.log(`🗑️ Brišem ${controlIndices.length} kontrola...\n`);

    // Briši unazad (stabilniji pristup)
    for (let i = controlIndices.length - 1; i >= 0; i--) {
      const idx = controlIndices[i];
      
      if (idx >= contentControls.items.length) {
        console.log(`⚠️ [${idx}] Indeks van opsega, preskačem`);
        continue;
      }

      const cc = contentControls.items[idx];
      const currentText = cc.text || "";
      const tag = cc.tag || "";
      const meta = parseTag(tag);

      if (!meta) {
        console.log(`⚠️ [${idx}] Kontrola više nije BA_FIELD, preskačem`);
        continue;
      }

      console.log(`🗑️ [${idx}] Brišem: ${meta.key}`);
      console.log(`    Tekst pre brisanja: "${currentText.substring(0, 60)}..."`);

      // Otključaj ako je zaključana
      if (cc.cannotDelete) {
        console.log(`    🔓 Otključavam kontrolu`);
        cc.cannotDelete = false;
      }

      // ⭐ KLJUČNA AKCIJA: Briši kontrolu, ZADRŽI TEKST
      cc.delete(false);
      removed++;
      
      console.log(`    ✅ Kontrola obrisana, tekst zadržan na istom mestu`);
    }

    await context.sync();
    console.log(`\n✅ Ukupno obrisano: ${removed} kontrola`);

    // Obriši XML state
    console.log("\n🔄 Brisanje XML state...");
    await deleteXMLState(context);
  });

  console.log("=".repeat(60));
  return removed;
}

/**
 * 🎯 GLAVNA FUNKCIJA - Entry point za Ribbon Command
 * Poziva se kada korisnik klikne dugme "Ukloni Kontrole"
 */
async function deleteAllContentControls(event) {
  console.log("\n🔴 deleteAllContentControls() pozvana iz Ribbon Command");
  console.log("⏰ Vreme: " + new Date().toLocaleTimeString());
  
  try {
    // FAZA 1: Mapiranje kontrola
    const mapping = await mapContentControls();
    
    if (mapping.controls.length === 0) {
      console.log("ℹ️ Nema BA_FIELD kontrola za brisanje");
      showNotification("Info", "Nisu pronađena aktivna polja u dokumentu.");
      event.completed();
      return;
    }

    // Pripremi podatke za dialog
    const dialogData = mapping.controls.map(ctrl => ({
      key: ctrl.key,
      text: ctrl.text,
      type: ctrl.type
    }));

    console.log(`\n💬 Prikazujem confirmation dialog sa ${dialogData.length} polja...`);

    // Prikaži confirmation dialog
    const dialogUrl = `https://baneandreev-byte.github.io/ba-word-addin/confirm-delete.html?controls=${encodeURIComponent(JSON.stringify(dialogData))}`;
    
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { 
        height: 60, 
        width: 45,
        displayInIframe: false 
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("❌ Greška pri otvaranju dijaloga:", result.error);
          showNotification("Greška", "Nije moguće otvoriti prozor za potvrdu.");
          event.completed();
          return;
        }

        const dialog = result.value;
        console.log("✅ Confirmation dialog otvoren");

        // Čekaj odgovor od dijaloga
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
          console.log("📨 Primljen odgovor od dijaloga:", arg.message);
          
          try {
            const response = JSON.parse(arg.message);
            
            dialog.close();
            console.log("🔒 Dialog zatvoren");

            if (response.confirmed) {
              console.log("✅ Korisnik potvrdio brisanje\n");
              
              // FAZA 2: Izvršavanje brisanja
              const controlIndices = mapping.controls.map(c => c.index);
              const removed = await deleteControlsByIndices(controlIndices);
              
              if (removed > 0) {
                const message = `Uklonjeno ${removed} aktivnih polja. Tekst zadržan u dokumentu.`;
                console.log(`\n✨ ${message}`);
                showNotification("Uspešno", message);
              }
            } else {
              console.log("❌ Korisnik otkazao brisanje");
              showNotification("Info", "Brisanje otkazano.");
            }
            
            event.completed();
            console.log("✅ Operacija završena\n");
            
          } catch (error) {
            console.error("❌ Greška pri obradi odgovora:", error);
            event.completed();
          }
        });

        // Handle dialog close
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          console.log("🔒 Dialog zatvoren (event):", arg.error);
          if (arg.error === 12006) {
            // User closed dialog
            console.log("ℹ️ Korisnik zatvorio dialog");
            showNotification("Info", "Brisanje otkazano.");
          }
          event.completed();
        });
      }
    );

  } catch (error) {
    console.error("❌ GREŠKA:", error);
    console.error("❌ Stack:", error.stack);
    showNotification("Greška", `Došlo je do greške: ${error.message}`);
    event.completed();
  }
}

/**
 * Prikaži notifikaciju korisniku (fallback - samo console log)
 */
function showNotification(title, message) {
  console.log(`📢 ${title}: ${message}`);
}

// ============================================
// REGISTRACIJA FUNKCIJA ZA OFFICE.JS
// ============================================
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("✅ Commands.js V47 loaded - Word detected");
    console.log("✅ Confirmation dialog implementiran");
    console.log("✅ Detaljno mapiranje pre brisanja");
    
    // Registruj funkcije za Ribbon Commands
    Office.actions.associate("deleteAllContentControls", deleteAllContentControls);
    
    console.log("✅ Ribbon Commands registered:");
    console.log("  - deleteAllContentControls (with confirmation)");
    console.log("=".repeat(60));
  }
});
