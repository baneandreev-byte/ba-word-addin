// ============================================
// PATCH ZA TASKPANE.JS - performDelete funkcija
// Zameni liniju 815-904 sa ovim kodom
// ============================================

/**
 * ⭐ POBOLJŠANA VERZIJA - insertText sa Replace location
 * Zamenjuje kontrolu sa tekstom bez dupliranja
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

        // ⭐ NOVA STRATEGIJA - ZAMENA kontrole sa tekstom
        // insertText sa Replace location briše kontrolu i ostavlja tekst
        if (currentText) {
          console.log(`    📝 Zamenjujem kontrolu sa tekstom`);
          cc.insertText(currentText, Word.InsertLocation.replace);
        } else {
          console.log(`    ⚠️ Kontrola je prazna, samo je brišem`);
          cc.delete(true);
        }
        
        removed++;
        console.log(`    ✅ Kontrola zamenjena tekstom`);
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

    setStatus(`Dokument očišćen: ${removed} polja uklonjeno.`, "info");
    closeDeleteModal();
    
  } catch (error) {
    console.error("❌ GREŠKA pri brisanju:", error);
    console.error("❌ Stack:", error.stack);
    setStatus("Greška pri brisanju polja. Vidi konzolu.", "error");
    closeDeleteModal();
  }
}
