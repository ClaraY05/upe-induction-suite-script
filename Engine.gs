function runCloningEngine(templateFolderId, newTerm, customDestFolderId) {
  const t0 = Date.now();
  const deadline = t0 + CONFIG.MAX_RUNTIME_MS;
  const warnings = [];
  const warn = msg => { console.warn(msg); warnings.push(msg); };

  newTerm = String(newTerm || "").trim();
  if (!newTerm) throw new Error("Term name cannot be empty.");
  if (PropertiesService.getScriptProperties().getProperty(JOB_KEY)) {
    throw new Error("A previous run is still finishing in the background. Wait for its " +
                    "\"_Provisioning report\" file, or run cancelPendingJob() in the editor.");
  }

  const templateFolder = DriveApp.getFolderById(templateFolderId);
  const destId = customDestFolderId || CONFIG.DESTINATION_PARENT_FOLDER_ID;
  const parentFolder = destId ? DriveApp.getFolderById(destId) : firstParentOrRoot(templateFolder);

  const tree = snapshotTree(templateFolder);
  const templateIds = new Set([...tree.folders.map(f => f.id), ...tree.files.map(f => f.id)]);
  if (templateIds.has(parentFolder.getId())) {
    throw new Error("Destination folder is inside the template folder. Choose a destination outside it.");
  }
  preflightPermissions(tree);

  const ctx = makeContext(newTerm, [...templateIds], [], warn);
  const byId = {};
  tree.files.forEach(f => { byId[f.id] = f; });

  const oldRoot = templateFolder.getName();
  let rootName = ctx.replaceKw(oldRoot);
  if (rootName === oldRoot) rootName = oldRoot.replace(/TEMPLATE/i, newTerm).trim();
  if (!rootName || rootName === oldRoot) rootName = `${oldRoot} - ${newTerm}`;

  const rootTarget = parentFolder.createFolder(rootName);
  const folderById = { [templateFolder.getId()]: rootTarget };
  tree.folders.forEach(f => {
    if (!f.parentId) return;
    folderById[f.id] = folderById[f.parentId].createFolder(ctx.replaceKw(f.name));
  });

  const sheetToForm = {};
  const oldPublished = {};
  tree.files.filter(f => f.mimeType === MimeType.GOOGLE_FORMS).forEach(f => {
    let form;
    try { form = FormApp.openById(f.id); }
    catch (e) { warn(`Couldn't open template form "${f.name}": ${e.message}`); return; }
    try { oldPublished[f.id] = publishedFormKey(form.getPublishedUrl()); } catch (e) {}
    try {
      if (form.getDestinationType() === FormApp.DestinationType.SPREADSHEET) {
        sheetToForm[form.getDestinationId()] = f;
      }
    } catch (e) {}
  });

  const idMap = {};
  const copied = [];
  let formsRelinked = 0;

  tree.files.filter(f => f.mimeType === MimeType.GOOGLE_SHEETS).forEach(f => {
    const newFile = copyOne(f, folderById[f.parentId], ctx);
    idMap[f.id] = newFile.getId();
    copied.push({ oldId: f.id, newId: newFile.getId(), mimeType: f.mimeType, name: f.name });

    const srcForm = sheetToForm[f.id] || null;
    if (!srcForm && !templateSheetHasForm(f.id)) return;

    const spawnedId = findSpawnedForm(newFile.getId(), srcForm ? srcForm.name : null);
    if (!spawnedId) {
      warn(`Couldn't find the form Google auto-created when copying "${f.name}". ` +
           `Check the template folder for a "Copy of ..." form and delete it.`);
      return;
    }
    ctx.assertWritable(spawnedId);
    const spawned = DriveApp.getFileById(spawnedId);
    const home = srcForm ? folderById[srcForm.parentId] : folderById[f.parentId];
    spawned.moveTo(home);
    const baseName = srcForm ? srcForm.name : spawned.getName().replace(/^Copy of /i, "");
    spawned.setName(ctx.replaceKw(baseName));

    if (srcForm) {
      idMap[srcForm.id] = spawnedId;
      formsRelinked++;
    } else {
      warn(`"${f.name}" is linked to a form outside the template; its auto-created copy was placed next to the sheet.`);
    }
    copied.push({ oldId: srcForm ? srcForm.id : null, newId: spawnedId, mimeType: MimeType.GOOGLE_FORMS, name: baseName });
  });

  const needsDestination = [];
  tree.files.filter(f => f.mimeType !== MimeType.GOOGLE_SHEETS).forEach(f => {
    if (idMap[f.id]) return;
    if (f.mimeType === MimeType.SHORTCUT) { warn(`Skipped shortcut "${f.name}".`); return; }
    const newFile = copyOne(f, folderById[f.parentId], ctx);
    idMap[f.id] = newFile.getId();
    copied.push({ oldId: f.id, newId: newFile.getId(), mimeType: f.mimeType, name: f.name });

    if (f.mimeType === MimeType.GOOGLE_FORMS) {
      const linkedSheet = Object.keys(sheetToForm).find(sid => sheetToForm[sid].id === f.id);
      if (linkedSheet) needsDestination.push({ formOld: f.id, sheetOld: linkedSheet });
    }
  });

  needsDestination.forEach(({ formOld, sheetOld }) => {
    const name = byId[formOld].name;
    if (!idMap[sheetOld]) { warn(`Form "${name}" responds to a sheet outside the template; the copy has no response sheet.`); return; }
    try {
      FormApp.openById(idMap[formOld]).setDestination(FormApp.DestinationType.SPREADSHEET, idMap[sheetOld]);
      formsRelinked++;
      warn(`Form "${name}" was linked via setDestination, so responses go to a NEW tab in its sheet.`);
    } catch (e) { warn(`Couldn't link form "${name}": ${e.message}`); }
  });

  Object.keys(idMap).forEach(o => ctx.addIdPair(o, idMap[o]));
  Object.keys(folderById).forEach(o => ctx.addIdPair(o, folderById[o].getId()));

  const task = kind => c => ({ kind, newId: c.newId, name: c.name,
                               oldPub: kind === "form" && c.oldId ? (oldPublished[c.oldId] || null) : null });
  const job = {
    newTerm,
    templateIds: [...templateIds],
    idPairs: ctx.idPairs,
    tasks: [
      ...copied.filter(c => c.mimeType === MimeType.GOOGLE_FORMS).map(task("form")),
      ...copied.filter(c => c.mimeType === MimeType.GOOGLE_SHEETS).map(task("sheet")),
      ...copied.filter(c => c.mimeType === MimeType.GOOGLE_DOCS).map(task("doc")),
    ],
    warnings,
    startedAt: t0,
    result: {
      newFolderName: rootName,
      newFolderUrl: rootTarget.getUrl(),
      newFolderId: rootTarget.getId(),
      filesCopied: copied.length,
      foldersCopied: tree.folders.length - 1,
      formsRelinked,
    },
  };
  return runTasks(job, deadline);
}

function runTasks(job, deadline) {
  const warn = msg => { console.warn(msg); job.warnings.push(msg); };
  const ctx = makeContext(job.newTerm, job.templateIds, job.idPairs, warn);

  while (job.tasks.length && Date.now() < deadline) {
    const t = job.tasks.shift();
    const label = { form: "Form", sheet: "Sheet", doc: "Doc" }[t.kind];
    safely(`${label} "${t.name}"`, () => {
      if (t.kind === "form") {
        const newPub = processForm(t.newId, ctx);
        if (t.oldPub && newPub) ctx.addIdPair(t.oldPub, newPub);
      } else if (t.kind === "sheet") {
        processSheet(t.newId, ctx);
      } else if (t.kind === "doc") {
        processDoc(t.newId, t.name, ctx);
      }
    }, warn);
  }
  job.idPairs = ctx.idPairs;

  if (job.tasks.length) {
    saveJob(job);
    scheduleContinuation();
  } else {
    finishJob(job);
  }
  return {
    ...job.result,
    warnings: job.warnings,
    pending: job.tasks.length,
    seconds: Math.round((Date.now() - job.startedAt) / 1000),
  };
}

function continueProvisioning() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    deleteContinuationTriggers();
    const fileId = PropertiesService.getScriptProperties().getProperty(JOB_KEY);
    if (!fileId) return;
    const job = JSON.parse(DriveApp.getFileById(fileId).getBlob().getDataAsString());
    runTasks(job, Date.now() + CONFIG.MAX_RUNTIME_MS);
  } finally {
    lock.releaseLock();
  }
}

function saveJob(job) {
  const props = PropertiesService.getScriptProperties();
  const json = JSON.stringify(job);
  const fileId = props.getProperty(JOB_KEY);
  if (fileId) {
    DriveApp.getFileById(fileId).setContent(json);
  } else {
    const f = DriveApp.getFolderById(job.result.newFolderId)
      .createFile("_provisioning-state.json", json, MimeType.PLAIN_TEXT);
    props.setProperty(JOB_KEY, f.getId());
  }
}

function finishJob(job) {
  const props = PropertiesService.getScriptProperties();
  const fileId = props.getProperty(JOB_KEY);
  if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
  props.deleteProperty(JOB_KEY);
  deleteContinuationTriggers();

  const secs = Math.round((Date.now() - job.startedAt) / 1000);
  const report =
    `Provisioning report for "${job.newTerm}"\n` +
    `Finished: ${new Date().toISOString()}  (${secs}s wall time)\n\n` +
    `Files copied: ${job.result.filesCopied}\n` +
    `Subfolders created: ${job.result.foldersCopied}\n` +
    `Forms linked to Sheets: ${job.result.formsRelinked}\n\n` +
    (job.warnings.length ? `Warnings (${job.warnings.length}):\n- ` + job.warnings.join("\n- ") : "No warnings.") + "\n";
  DriveApp.getFolderById(job.result.newFolderId)
    .createFile("_Provisioning report.txt", report, MimeType.PLAIN_TEXT);
}

function scheduleContinuation() {
  deleteContinuationTriggers();
  ScriptApp.newTrigger("continueProvisioning").timeBased().after(60 * 1000).create();
}

function deleteContinuationTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "continueProvisioning") ScriptApp.deleteTrigger(t);
  });
}

function preflightPermissions(tree) {
  const missing = [];
  const probe = (label, fn) => {
    try { fn(); } catch (e) { if (/permission|authoriz/i.test(e.message)) missing.push(label); }
  };
  const form = tree.files.find(f => f.mimeType === MimeType.GOOGLE_FORMS);
  const doc = tree.files.find(f => f.mimeType === MimeType.GOOGLE_DOCS);
  const sheet = tree.files.find(f => f.mimeType === MimeType.GOOGLE_SHEETS);
  if (form) probe("Forms", () => FormApp.openById(form.id).getTitle());
  if (doc) probe("Docs", () => DocumentApp.openById(doc.id).getName());
  if (sheet) probe("Sheets", () => SpreadsheetApp.openById(sheet.id).getName());
  probe("Triggers", () => ScriptApp.getProjectTriggers());
  if (missing.length) {
    throw new Error(
      `Missing authorization for: ${missing.join(", ")}. Nothing was copied. ` +
      `In the editor: confirm appsscript.json lists the scopes, run authorizeScript, ` +
      `accept every permission, then Deploy > Manage deployments > New version.`
    );
  }
}