/* global Office, Word */

// ============================================
// VERZIJA: 2025-02-08 - V45 (RIBBON COMMANDS)
// commands.js - Ribbon Command Functions
// ============================================
console.log("🔧 BA Word Add-in Commands VERZIJA: 2025-02-08 - V45");
console.log("✅ Ribbon Command za brisanje content controls");

/**
 * Parse BA_FIELD tag to extract metadata
 * Format: BA_FIELD|key=NAZIV|type=text|format=text:auto
 */
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
    // Ne throw-uj grešku - XML state nije kritičan
  }
}

/**
 * ⭐ Glavna funkcija - Briše sve BA_FIELD content control-e iz dokumenta
 * Zadržava tekst, briše kontrole i XML state
 * POZIVA SE IZ RIBBON COMMAND DUGMETA
 */
async function deleteAllContentControls(event) {
  console.log("🔴 deleteAllContentControls() pozvana iz Ribbon Command");
  
  try {
    let removed = 0;
    let xmlDeleted = false;

    await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();

      const totalControls = contentControls.items.length;
      console.log(`📊 Pronađeno ${totalControls} content controls u dokumentu`);

      if (totalControls === 0) {
        console.log("ℹ️ Nema content control-a za brisanje");
        event.completed();
        return;
      }

      // FAZA 1: Učitaj properties za sve kontrole
      for (const cc of contentControls.items) {
        cc.load("tag,text,cannotDelete");
      }
      await context.sync();
      console.log("✅ Učitane properties za sve kontrole");

      // FAZA 2: Procesuj samo BA_FIELD kontrole - unazad
      const toDelete = [];
      
      for (let i = contentControls.items.length - 1; i >= 0; i--) {
        const cc = contentControls.items[i];
        const meta = parseTag(cc.tag || "");
        
        // Preskači ako nije BA_FIELD
        if (!meta) {
          console.log(`⏭️ Preskačem kontrolu [${i}]: nije BA_FIELD format`);
          continue;
        }

        console.log(`🔍 Procesiranje kontrole [${i}]: ${meta.key}`);

        // Otključaj ako je zaključana
        if (cc.cannotDelete) {
          console.log(`  🔓 Otključavanje kontrole: ${meta.key}`);
          cc.cannotDelete = false;
        }

        // Sačuvaj tekst
        const currentText = cc.text || "";
        console.log(`  📝 Tekst u kontroli: "${currentText}"`);

        // ⭐ KRITIČNA AKCIJA: Obriši kontrolu, ZADRŽI TEKST
        // delete(false) = zadrži sadržaj u dokumentu
        cc.delete(false);
        toDelete.push(meta.key);
        removed++;
        
        console.log(`  ✅ Kontrola "${meta.key}" obrisana (tekst zadržan)`);
      }

      await context.sync();
      console.log(`✅ Obrisano ${removed} BA_FIELD kontrola`);

      // FAZA 3: Obriši XML state ako postoji
      await deleteXMLState(context);
      xmlDeleted = true;
    });

    // Prikaži rezultat korisniku
    if (removed > 0) {
      const message = xmlDeleted 
        ? `Uklonjeno ${removed} kontrola. Tekst sačuvan, plugin podaci obrisani.`
        : `Uklonjeno ${removed} kontrola. Tekst sačuvan.`;
      
      console.log(`✨ ${message}`);
      
      // Notification preko Office.ui
      showRibbonNotification(
        "Uspešno", 
        message
      );
    } else {
      console.log("ℹ️ Nisu pronađene BA_FIELD kontrole");
      showRibbonNotification(
        "Info", 
        "Nisu pronađene BiroA kontrole u dokumentu."
      );
    }

  } catch (error) {
    console.error("❌ Greška pri brisanju content control-a:", error);
    console.error("❌ Stack:", error.stack);
    
    showRibbonNotification(
      "Greška", 
      `Došlo je do greške: ${error.message}`
    );
  }

  // ⚠️ OBAVEZNO za ExecuteFunction akcije
  event.completed();
}

/**
 * Prikaz notifikacije korisniku
 * Koristi Office.addin.showAsTaskpane() ili message bar
 */
function showRibbonNotification(title, message) {
  try {
    // Office.addin API za notifikacije (Office 2016+)
    if (Office.context.ui && Office.context.ui.displayDialogAsync) {
      // Prikaži kao info bar u dokumentu
      console.log(`📢 ${title}: ${message}`);
      
      // Alternativno: Možemo koristiti dialog za bolje iskustvo
      // Ali za sada samo logujemo - Office.addin.showAsTaskpane zahteva HTML
    } else {
      // Fallback - samo console log
      console.log(`📢 ${title}: ${message}`);
    }
  } catch (error) {
    console.error("⚠️ Greška pri prikazu notifikacije:", error);
  }
}

/**
 * ⭐ NAPREDNA VERZIJA - Sa confirmation dijalogom
 * Može se implementirati kasnije ako je potrebno
 */
async function deleteContentControlsWithConfirm(event) {
  try {
    // Prvo proveri koliko ima kontrola
    let controlCount = 0;
    await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();
      
      for (const cc of contentControls.items) {
        cc.load("tag");
      }
      await context.sync();
      
      // Prebroj samo BA_FIELD kontrole
      for (const cc of contentControls.items) {
        if (parseTag(cc.tag)) {
          controlCount++;
        }
      }
    });
    
    if (controlCount === 0) {
      showRibbonNotification("Info", "Nema BiroA kontrola za brisanje");
      event.completed();
      return;
    }
    
    // Otvori confirmation dialog
    Office.context.ui.displayDialogAsync(
      'https://baneandreev-byte.github.io/ba-word-addin/confirm-delete.html?count=' + controlCount,
      { height: 30, width: 40 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
            const response = JSON.parse(arg.message);
            
            if (response.confirmed) {
              // Korisnik je potvrdio - pozovi glavnu funkciju
              await deleteAllContentControls(event);
            } else {
              console.log("ℹ️ Korisnik je otkazao brisanje");
              event.completed();
            }
            
            dialog.close();
          });
        } else {
          console.error("❌ Greška pri otvaranju dijaloga:", result.error);
          event.completed();
        }
      }
    );
  } catch (error) {
    console.error("❌ Greška:", error);
    event.completed();
  }
}

// ============================================
// REGISTRACIJA FUNKCIJA ZA OFFICE.JS
// ============================================
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("✅ Commands.js loaded - Word detected");
    
    // Registruj funkcije za Ribbon Commands
    Office.actions.associate("deleteAllContentControls", deleteAllContentControls);
    
    console.log("✅ Ribbon Commands registered:");
    console.log("  - deleteAllContentControls");
  }
});
