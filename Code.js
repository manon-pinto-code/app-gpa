// ============================================================
// 📋 GPA AUTO-FORWARD — RPI Menuiserie (VERSION TEST)
// ============================================================
// 🧪 TEST CONFIG :
//   - Expéditrice (commercial) : antoine.bartra@rpimenuiserie.com
//   - Destination (centralisée) : rpimenuiserie33@gmail.com
//
// ⚠️  Une fois validé, modifier DESTINATION_EMAIL avec
//     l'adresse de production (sav.promotion@rpimenuiserie.com)
//     et déployer sur chaque boîte commerciale.
// ============================================================

const CONFIG = {
  DESTINATION_EMAIL: "sav.promotion@rpimenuiserie.com",

  LABEL_TO_PROCESS: "📋 GPA",
  LABEL_DONE: "✅ GPA Transféré",

  SUBJECT_PREFIX: "[GPA]",
  MAX_PER_RUN: 20,

  ENABLE_SHEET_LOG: true,
  LOG_SHEET_NAME: "LOG_GPA_Forward",
};


// ── FONCTION PRINCIPALE (toutes les 5 min) ──────────────────
function processGPA() {
  const labelToProcess = GmailApp.getUserLabelByName(CONFIG.LABEL_TO_PROCESS);
  const labelDone = getOrCreateLabel_(CONFIG.LABEL_DONE);

  if (!labelToProcess) {
    GmailApp.createLabel(CONFIG.LABEL_TO_PROCESS);
    Logger.log("Libellé '" + CONFIG.LABEL_TO_PROCESS + "' créé. En attente de mails.");
    return;
  }

  const threads = labelToProcess.getThreads(0, CONFIG.MAX_PER_RUN);

  if (threads.length === 0) {
    Logger.log("Aucun mail GPA à traiter.");
    return;
  }

  Logger.log(threads.length + " thread(s) GPA à traiter.");

  let successCount = 0;
  let errorCount = 0;

  for (const thread of threads) {
    try {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];

      lastMessage.forward(CONFIG.DESTINATION_EMAIL, {
        subject: CONFIG.SUBJECT_PREFIX + " " + lastMessage.getSubject(),
      });

      thread.removeLabel(labelToProcess);
      thread.addLabel(labelDone);

      const logEntry = {
        date: new Date(),
        from: lastMessage.getFrom(),
        subject: lastMessage.getSubject(),
        commercial: Session.getActiveUser().getEmail(),
        status: "OK",
      };

      Logger.log("✅ Transféré : " + logEntry.subject + " (de: " + logEntry.from + ")");

      if (CONFIG.ENABLE_SHEET_LOG) {
        logToSheet_(logEntry);
      }

      successCount++;

    } catch (error) {
      errorCount++;
      Logger.log("❌ Erreur : " + thread.getFirstMessageSubject() + " — " + error.message);

      if (CONFIG.ENABLE_SHEET_LOG) {
        logToSheet_({
          date: new Date(),
          from: "N/A",
          subject: thread.getFirstMessageSubject(),
          commercial: Session.getActiveUser().getEmail(),
          status: "ERREUR: " + error.message,
        });
      }
    }
  }

  Logger.log("Terminé. " + successCount + " transféré(s), " + errorCount + " erreur(s).");
}


// ── UTILITAIRES ─────────────────────────────────────────────

function getOrCreateLabel_(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log("Libellé '" + labelName + "' créé.");
  }
  return label;
}

function logToSheet_(entry) {
  try {
    let ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {
      ss = null;
    }

    if (!ss) {
      ss = createLogSpreadsheet_();
    }

    let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
      sheet.appendRow(["Date", "Expéditeur original", "Objet", "Commercial", "Statut"]);
      sheet.getRange("1:1").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      entry.date,
      entry.from,
      entry.subject,
      entry.commercial,
      entry.status,
    ]);

  } catch (e) {
    Logger.log("⚠️ Log Sheet échoué (non-bloquant) : " + e.message);
  }
}

function createLogSpreadsheet_() {
  const ss = SpreadsheetApp.create("🧪 GPA Forward — Logs TEST");
  Logger.log("Sheet de logs créé : " + ss.getUrl());
  return ss;
}


// ── SETUP (exécuter UNE FOIS) ───────────────────────────────

function setup() {
  getOrCreateLabel_(CONFIG.LABEL_TO_PROCESS);
  getOrCreateLabel_(CONFIG.LABEL_DONE);
  Logger.log("✅ Libellés créés.");

  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "processGPA") {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger("processGPA")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("✅ Setup terminé.");
  Logger.log("📬 Boîte source : antoine.bartra@rpimenuiserie.com");
  Logger.log("📥 Destination  : rpimenuiserie33@gmail.com");
  Logger.log("⏱️  Trigger     : toutes les 5 minutes");
  Logger.log("ℹ️  Classez vos mails GPA dans le libellé '" + CONFIG.LABEL_TO_PROCESS + "'.");
}


// ── DÉSINSTALLATION ─────────────────────────────────────────

function uninstall() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "processGPA") {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  }
  Logger.log("✅ " + count + " trigger(s) supprimé(s). Script désactivé.");
}