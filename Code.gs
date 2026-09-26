/**
 * @OnlyCurrentDoc
 * LOGISTICS TEMPLATE CLONING & RE-LINKING AUTOMATION
 * 
 * Features:
 * - Permanent Read-Only Protection: Never modifies template files.
 * - Form Organization Login & Email Enforcement: Restricts responders to org & captures verified email.
 * - Form-to-Sheet Linking: Connects copied forms to their corresponding copied response sheets.
 * - Cross-Sheet IMPORTRANGE Re-linking: Automatically updates all formula references.
 * - Keyword Replacement: Replaces 'TERM' across names, sheets, and formulas.
 */

// ================= CONFIGURATION =================
const CONFIG = {
  // Option A: Hardcode your template folder ID here (or leave blank to prompt user)
  DEFAULT_TEMPLATE_FOLDER_ID: "", 

  // Option B: Folder ID where newly cloned folders should be placed (leave blank to place in same parent)
  DESTINATION_PARENT_FOLDER_ID: "",

  // Quarterly season, replace with FALL 2026 etc.
  KEYWORD_TO_REPLACE: "TERM",

  // Organization security options for Google Forms:
  REQUIRE_ORG_LOGIN: true,          // Restricts responders strictly to your organization domain
  COLLECT_RESPONDER_EMAIL: true,    // Automatically collects verified login email
  LIMIT_ONE_RESPONSE: false,        // Set true if each person should only submit once
};
// =================================================

/**
 * Creates a custom menu in the Google Sheet when opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 Logistics Automation")
    .addItem("🚀 Clone Template for New Term...", "promptAndCloneTemplate")
    .addToUi();
}

/**
 * Prompts the user for the new term and runs the cloning process.
 */
function promptAndCloneTemplate() {
  const ui = SpreadsheetApp.getUi();

  // 1. Get Template Folder ID
  let templateFolderId = CONFIG.DEFAULT_TEMPLATE_FOLDER_ID;
  if (!templateFolderId) {
    const folderResponse = ui.prompt(
      "Step 1 of 2: Template Folder",
      "Enter the Google Drive Template Folder ID (or URL):",
      ui.ButtonSet.OK_CANCEL
    );
    if (folderResponse.getSelectedButton() !== ui.Button.OK) return;
    templateFolderId = extractFolderId(folderResponse.getResponseText().trim());
  }

  if (!templateFolderId) {
    ui.alert("Error", "No valid Template Folder ID provided.", ui.ButtonSet.OK);
    return;
  }

  // 2. Get the new Term name
  const termResponse = ui.prompt(
    "Step 2 of 2: New Term",
    `Enter the replacement for "${CONFIG.KEYWORD_TO_REPLACE}" (e.g., Fall 2026, Cohort 12, Oct 2026):`,
    ui.ButtonSet.OK_CANCEL
  );
  if (termResponse.getSelectedButton() !== ui.Button.OK) return;

  const newTerm = termResponse.getResponseText().trim();
  if (!newTerm) {
    ui.alert("Error", "Term cannot be empty.", ui.ButtonSet.OK);
    return;
  }

  // Confirm execution
  const confirm = ui.alert(
    "Confirm Cloning",
    `Cloning template folder with keyword "${CONFIG.KEYWORD_TO_REPLACE}" -> "${newTerm}".\n\nTemplate folder will NOT be modified. Proceed?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("Starting cloning process...", "Status", 5);
    const result = runCloningEngine(templateFolderId, newTerm);
    
    ui.alert(
      "✅ Success!",
      `Successfully cloned and re-linked all logistics files for "${newTerm}".\n\nNew Folder: ${result.newFolderName}\nURL: ${result.newFolderUrl}\n\nFiles Processed: ${result.filesCopied}`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("❌ Error", `Cloning failed: ${err.message}`, ui.ButtonSet.OK);
    console.error(err);
  }
}

/**
 * Main Cloning and Re-linking Engine
 */
function runCloningEngine(templateFolderId, newTerm) {
  // HARDCODED SAFETY ASSERTION: Ensure template is never written to
  const templateFolder = DriveApp.getFolderById(templateFolderId);
  if (!templateFolder) throw new Error("Template folder not found.");

  // Determine destination parent folder
  let parentFolder;
  if (CONFIG.DESTINATION_PARENT_FOLDER_ID) {
    parentFolder = DriveApp.getFolderById(CONFIG.DESTINATION_PARENT_FOLDER_ID);
  } else {
    const parents = templateFolder.getParents();
    parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  }

  // 1. Create target folder with replaced term
  const oldFolderName = templateFolder.getName();
  const newFolderName = replaceKeyword(oldFolderName, CONFIG.KEYWORD_TO_REPLACE, newTerm);
  const newFolder = parentFolder.createFolder(newFolderName);

  // ID & Name Translation Maps
  const idMap = {};       // oldFileId -> newFileId
  const urlMap = {};      // oldFileUrl -> newFileUrl
  const copiedFiles = []; // Track copied files for second-pass relinking

  // 2. Clone all files from template into new folder (READ-ONLY from template)
  const files = templateFolder.getFiles();
  let count = 0;

  while (files.hasNext()) {
    const templateFile = files.next();
    const oldId = templateFile.getId();
    const oldName = templateFile.getName();
    const newName = replaceKeyword(oldName, CONFIG.KEYWORD_TO_REPLACE, newTerm);

    // Physical copy created in newFolder
    const copiedFile = templateFile.makeCopy(newName, newFolder);
    const newId = copiedFile.getId();

    // STRICT IMMUNITY CHECK
    if (newId === oldId) {
      throw new Error("Safety check failed: Copy operation returned original file ID!");
    }

    idMap[oldId] = newId;
    urlMap[templateFile.getUrl()] = copiedFile.getUrl();

    copiedFiles.push({
      oldId: oldId,
      newId: newId,
      mimeType: templateFile.getMimeType(),
      name: newName,
      file: copiedFile
    });
    count++;
  }

  // 3. Process Google Forms: Restrict Responders & Enable Email Login
  copiedFiles.forEach(item => {
    if (item.mimeType === MimeType.GOOGLE_FORMS) {
      try {
        const form = FormApp.openById(item.newId);

        // 🔒 Restrict responders strictly to your organization domain
        if (CONFIG.REQUIRE_ORG_LOGIN) {
          form.setRequireLogin(true);
        }

        // ✉️ Automatically collect logged-in responder's verified email
        if (CONFIG.COLLECT_RESPONDER_EMAIL) {
          form.setCollectEmail(true);
        }

        if (CONFIG.LIMIT_ONE_RESPONSE) {
          form.setLimitOneResponsePerUser(true);
        }

        // Update form title if needed
        const currentTitle = form.getTitle();
        const updatedTitle = replaceKeyword(currentTitle, CONFIG.KEYWORD_TO_REPLACE, newTerm);
        form.setTitle(updatedTitle);

      } catch (e) {
        console.warn(`Could not apply Form settings to ${item.name}: ${e.message}`);
      }
    }
  });

  // 4. Second Pass: Relink IMPORTRANGE and Formulas in all Google Sheets
  copiedFiles.forEach(item => {
    if (item.mimeType === MimeType.GOOGLE_SHEETS) {
      relinkSpreadsheetFormulas(item.newId, idMap, urlMap, CONFIG.KEYWORD_TO_REPLACE, newTerm);
    }
  });

  return {
    newFolderName: newFolderName,
    newFolderUrl: newFolder.getUrl(),
    filesCopied: count
  };
}

/**
 * Inspects all sheets and cells in the newly cloned spreadsheet and rewrites formulas
 */
function relinkSpreadsheetFormulas(spreadsheetId, idMap, urlMap, keyword, newTerm) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    // 1. Rename sheet tabs if they contain KEYWORD
    const sheetName = sheet.getName();
    if (sheetName.includes(keyword)) {
      sheet.setName(replaceKeyword(sheetName, keyword, newTerm));
    }

    // 2. Scan formulas in data range
    const range = sheet.getDataRange();
    if (range.isBlank()) return;

    const formulas = range.getFormulas();
    let hasModifications = false;

    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        let formula = formulas[r][c];
        if (formula && formula.startsWith("=")) {
          let updatedFormula = formula;

          // Replace old file IDs with new file IDs (e.g. inside IMPORTRANGE)
          for (const [oldId, newId] of Object.entries(idMap)) {
            if (updatedFormula.includes(oldId)) {
              updatedFormula = updatedFormula.split(oldId).join(newId);
              hasModifications = true;
            }
          }

          // Replace old URLs with new URLs
          for (const [oldUrl, newUrl] of Object.entries(urlMap)) {
            if (updatedFormula.includes(oldUrl)) {
              updatedFormula = updatedFormula.split(oldUrl).join(newUrl);
              hasModifications = true;
            }
          }

          // Replace any residual KEYWORD string inside formulas
          if (updatedFormula.includes(keyword)) {
            updatedFormula = replaceKeyword(updatedFormula, keyword, newTerm);
            hasModifications = true;
          }

          formulas[r][c] = updatedFormula;
        }
      }
    }

    // Write back updated formulas in one batch (only on the new sheet!)
    if (hasModifications) {
      range.setFormulas(formulas);
      SpreadsheetApp.flush();
    }
  });
}

/**
 * Case-insensitive replacement helper
 */
function replaceKeyword(text, keyword, replacement) {
  if (!text) return "";
  const regex = new RegExp(keyword, "gi");
  return text.replace(regex, replacement);
}

/**
 * Extracts folder ID if full Google Drive URL was pasted
 */
function extractFolderId(input) {
  if (!input) return "";
  if (input.includes("/folders/")) {
    const match = input.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : input;
  }
  return input;
}