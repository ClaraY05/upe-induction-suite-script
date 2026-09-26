function processDoc(docId, name, ctx) {
  ctx.assertWritable(docId);
  const doc = DocumentApp.openById(docId);
  const stats = { links: 0, chips: 0 };

  docSections(doc).forEach(sec => {
    sec.replaceText(ctx.re2Pattern, ctx.newTerm);
    if (ctx.idTest) remapDocLinks(sec, ctx, stats);
  });
  doc.saveAndClose();

  if (stats.chips) {
    ctx.warn(`Doc "${name}": ${stats.chips} smart chip(s) still point at template files. ` +
             `Apps Script can't edit Docs chips, so replace them by hand.`);
  }
}

function docSections(doc) {
  const out = [];
  const add = x => { if (x) out.push(x); };
  if (typeof doc.getTabs === "function") {
    const walk = tabs => tabs.forEach(tab => {
      try {
        const dt = tab.asDocumentTab();
        add(dt.getBody()); add(dt.getHeader()); add(dt.getFooter());
      } catch (e) {}
      walk(tab.getChildTabs());
    });
    walk(doc.getTabs());
  }
  if (!out.length) { add(doc.getBody()); add(doc.getHeader()); add(doc.getFooter()); }
  return out;
}

function remapDocLinks(el, ctx, stats) {
  const type = el.getType();
  if (type === DocumentApp.ElementType.TEXT) {
    const t = el.asText();
    const len = t.getText().length;
    const idx = t.getTextAttributeIndices();
    for (let i = 0; i < idx.length; i++) {
      const start = idx[i];
      const end = (i + 1 < idx.length ? idx[i + 1] : len) - 1;
      if (end < start) continue;
      const url = t.getLinkUrl(start);
      if (!url) continue;
      const nu = ctx.remap(url);
      if (nu !== url) { t.setLinkUrl(start, end, nu); stats.links++; }
    }
    return;
  }
  if (type === DocumentApp.ElementType.RICH_LINK) {
    const u = el.asRichLink().getUrl();
    if (u && ctx.idTest.test(u)) stats.chips++;
    return;
  }
  if (typeof el.getNumChildren === "function") {
    for (let i = 0; i < el.getNumChildren(); i++) remapDocLinks(el.getChild(i), ctx, stats);
  }
}