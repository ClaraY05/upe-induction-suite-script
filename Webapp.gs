function executeFromWeb(folderId, destFolderId, newTerm) {
  const t = extractFolderId(folderId) || CONFIG.DEFAULT_TEMPLATE_FOLDER_ID;
  const d = extractFolderId(destFolderId) || CONFIG.DESTINATION_PARENT_FOLDER_ID;
  return runCloningEngine(t, newTerm, d);
}

function doGet() {
  const t = HtmlService.createTemplateFromFile("Index");
  t.defaultTemplateId = CONFIG.DEFAULT_TEMPLATE_FOLDER_ID;
  t.defaultDestId = CONFIG.DESTINATION_PARENT_FOLDER_ID;
  return t.evaluate()
    .setTitle("UPE I&M Suite Script")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}