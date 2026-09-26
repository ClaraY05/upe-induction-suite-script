function makeContext(newTerm, templateIdList, idPairs, warn) {
  const templateIds = new Set(templateIdList);
  const esc = escapeRegex(CONFIG.KEYWORD_TO_REPLACE);
  const pattern = CONFIG.WHOLE_WORD ? `\\b${esc}\\b` : esc;
  const flags = CONFIG.MATCH_CASE ? "" : "i";

  const ctx = {
    newTerm, templateIds, warn, pattern,
    re2Pattern: CONFIG.MATCH_CASE ? pattern : "(?i)" + pattern,
    idPairs: (idPairs || []).slice(),
    idTest: null,

    hasKw: s => !!s && new RegExp(pattern, flags).test(s),
    replaceKw: s => (s ? s.replace(new RegExp(pattern, flags + "g"), () => newTerm) : s),

    remap(s) {
      if (!s || !ctx.idTest || !ctx.idTest.test(s)) return s;
      let out = s;
      for (const [o, n] of ctx.idPairs) if (out.indexOf(o) !== -1) out = out.split(o).join(n);
      return out;
    },
    addIdPair(o, n) {
      if (o && n && o !== n && !ctx.idPairs.some(p => p[0] === o)) { ctx.idPairs.push([o, n]); ctx.finalizeIds(); }
    },
    finalizeIds() {
      ctx.idPairs.sort((a, b) => b[0].length - a[0].length);
      ctx.idTest = ctx.idPairs.length
        ? new RegExp(ctx.idPairs.map(p => escapeRegex(p[0])).join("|"))
        : null;
    },
    assertWritable(id) {
      if (templateIds.has(id)) throw new Error(`SAFETY: refusing to modify template file ${id}`);
    },
  };
  ctx.finalizeIds();
  return ctx;
}

function snapshotTree(root) {
  const folders = [{ id: root.getId(), name: root.getName(), parentId: null }];
  const files = [];
  const queue = [root];
  while (queue.length) {
    const folder = queue.shift();
    const fid = folder.getId();
    const fi = folder.getFiles();
    while (fi.hasNext()) {
      const f = fi.next();
      files.push({ id: f.getId(), name: f.getName(), mimeType: f.getMimeType(), parentId: fid });
    }
    const si = folder.getFolders();
    while (si.hasNext()) {
      const s = si.next();
      folders.push({ id: s.getId(), name: s.getName(), parentId: fid });
      queue.push(s);
    }
  }
  return { folders, files };
}

function copyOne(f, targetFolder, ctx) {
  const copy = DriveApp.getFileById(f.id).makeCopy(ctx.replaceKw(f.name), targetFolder);
  if (copy.getId() === f.id) throw new Error("SAFETY: makeCopy returned the template file's own ID.");
  return copy;
}

function templateSheetHasForm(sheetId) {
  try { return !!SpreadsheetApp.openById(sheetId).getFormUrl(); } catch (e) { return false; }
}

function findSpawnedForm(newSheetId, srcFormName) {
  for (let i = 0; i < CONFIG.SPAWNED_FORM_RETRIES; i++) {
    const url = SpreadsheetApp.openById(newSheetId).getFormUrl();
    if (url) {
      try { return FormApp.openByUrl(url).getId(); } catch (e) {}
      const m = url.match(/\/forms\/d\/(?!e\/)([\w-]{20,})/);
      if (m) return m[1];
      break;
    }
    Utilities.sleep(CONFIG.SPAWNED_FORM_WAIT_MS);
  }
  if (!srcFormName) return null;
  const since = Utilities.formatDate(new Date(Date.now() - 15 * 60 * 1000), "UTC", "yyyy-MM-dd'T'HH:mm:ss");
  const safeName = srcFormName.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
  const it = DriveApp.searchFiles(
    `title contains '${safeName}' and mimeType = '${MimeType.GOOGLE_FORMS}' and trashed = false and modifiedDate > '${since}'`
  );
  while (it.hasNext()) {
    const f = it.next();
    try { if (FormApp.openById(f.getId()).getDestinationId() === newSheetId) return f.getId(); } catch (e) {}
  }
  return null;
}

function publishedFormKey(url) {
  const m = url && url.match(/\/forms\/d\/e\/([\w-]+)/);
  return m ? m[1] : null;
}

function safely(label, fn, warn) {
  try { fn(); } catch (e) { warn(`${label}: ${e.message}`); }
}

function firstParentOrRoot(folder) {
  const p = folder.getParents();
  return p.hasNext() ? p.next() : DriveApp.getRootFolder();
}

function escapeRegex(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractFolderId(input) {
  if (!input) return "";
  const m = input.match(/\/folders\/([\w-]+)/) || input.match(/[?&]id=([\w-]+)/);
  return m ? m[1] : input.trim();
}