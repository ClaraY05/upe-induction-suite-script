function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("🎓 Induction Setup")
      .addItem("🚀 Provision New Quarter...", "promptAndCloneTemplate")
      .addToUi();
  } catch (e) {}
}

function authorizeScript() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
  console.log("Status: " + ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationStatus());
}

function promptAndCloneTemplate() {
  const ui = SpreadsheetApp.getUi();

  let templateFolderId = CONFIG.DEFAULT_TEMPLATE_FOLDER_ID;
  if (!templateFolderId) {
    const r = ui.prompt("Step 1 of 2: Template Folder", "Enter the Template Folder ID or URL:", ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    templateFolderId = extractFolderId(r.getResponseText().trim());
  }

  const termResponse = ui.prompt(
    "Step 2 of 2: New Quarter / Term Name",
    `Enter the replacement for "${CONFIG.KEYWORD_TO_REPLACE}" (e.g., Fall 2026):`,
    ui.ButtonSet.OK_CANCEL
  );
  if (termResponse.getSelectedButton() !== ui.Button.OK) return;
  const newTerm = termResponse.getResponseText().trim();
  if (!newTerm) { ui.alert("Error", "Term name cannot be empty.", ui.ButtonSet.OK); return; }

  const confirm = ui.alert(
    "Confirm Provisioning",
    `Cloning template tree for "${newTerm}".\nDestination: ${CONFIG.DESTINATION_PARENT_FOLDER_ID}\n\nProceed?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("Copying files and linking forms...", "Status", 10);
    const res = runCloningEngine(templateFolderId, newTerm, CONFIG.DESTINATION_PARENT_FOLDER_ID);
    const w = res.warnings.length
      ? `\n\n⚠️ ${res.warnings.length} warning(s):\n• ` + res.warnings.slice(0, 10).join("\n• ") +
        (res.warnings.length > 10 ? "\n(see Executions log for the rest)" : "")
      : "";
    ui.alert(
      "✅ Provisioning Complete",
      `📁 ${res.newFolderName}\n🔗 ${res.newFolderUrl}\n` +
      `📄 Files: ${res.filesCopied}   📂 Subfolders: ${res.foldersCopied}\n` +
      `🔗 Forms linked to Sheets: ${res.formsRelinked}   ⏱ ${res.seconds}s` +
      (res.pending ? `\n\n⏳ ${res.pending} item(s) are still processing in the background. ` +
                     `A "_Provisioning report" file will appear in the new folder when done.` : "") + w,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("❌ Error", `Provisioning failed: ${err.message}`, ui.ButtonSet.OK);
    console.error(err);
  }
}

function runAll() {
  const NEW_TERM = "Fall 2026";   // set this before running

  const status = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationStatus();
  console.log("Authorization status: " + status);
  if (String(status) !== "NOT_REQUIRED") {
    throw new Error("Not fully authorized. Run authorizeScript first and accept every permission.");
  }

  const tree = snapshotTree(DriveApp.getFolderById(CONFIG.DEFAULT_TEMPLATE_FOLDER_ID));
  const bloat = tree.files.filter(f => /^Copy of /i.test(f.name));
  if (bloat.length) {
    bloat.forEach(f => console.log(`BLOAT: ${f.name} -> https://drive.google.com/open?id=${f.id}`));
    throw new Error(`${bloat.length} "Copy of..." file(s) in the template. Delete or rename them and run again.`);
  }

  const res = runCloningEngine(CONFIG.DEFAULT_TEMPLATE_FOLDER_ID, NEW_TERM, CONFIG.DESTINATION_PARENT_FOLDER_ID);
  console.log(`Folder: ${res.newFolderUrl}`);
  console.log(`Files: ${res.filesCopied} | Subfolders: ${res.foldersCopied} | Forms linked: ${res.formsRelinked} | ${res.seconds}s`);
  if (res.pending) console.log(`${res.pending} item(s) will finish in the background; watch for "_Provisioning report.txt".`);
  res.warnings.forEach(w => console.warn(w));
}

function checkScopes() {
  console.log("Status: " + ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationStatus());
  const tree = snapshotTree(DriveApp.getFolderById(CONFIG.DEFAULT_TEMPLATE_FOLDER_ID));
  const form = tree.files.find(f => f.mimeType === MimeType.GOOGLE_FORMS);
  const doc = tree.files.find(f => f.mimeType === MimeType.GOOGLE_DOCS);
  const probe = (label, fn) => { try { fn(); console.log(label + ": OK"); } catch (e) { console.log(label + ": FAILED - " + e.message); } };
  if (form) probe("Forms", () => FormApp.openById(form.id).getTitle());
  if (doc) probe("Docs", () => DocumentApp.openById(doc.id).getName());
  probe("Triggers", () => ScriptApp.getProjectTriggers());
}

function listTemplateBloat() {
  const tree = snapshotTree(DriveApp.getFolderById(CONFIG.DEFAULT_TEMPLATE_FOLDER_ID));
  const suspects = tree.files.filter(f => /^Copy of /i.test(f.name));
  suspects.forEach(f => console.log(`${f.name}  ->  https://drive.google.com/open?id=${f.id}`));
  console.log(`${suspects.length} suspect file(s). Review and delete by hand.`);
}

function cancelPendingJob() {
  PropertiesService.getScriptProperties().deleteProperty(JOB_KEY);
  deleteContinuationTriggers();
  console.log("Pending job cleared.");
}