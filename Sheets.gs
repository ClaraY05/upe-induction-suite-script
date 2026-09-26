function processSheet(ssId, ctx) {
  ctx.assertWritable(ssId);
  const ss = SpreadsheetApp.openById(ssId);

  const chipCells = rewriteChips(ssId, ctx);

  if (ctx.idTest) {
    ss.getSheets().forEach(sheet => {
      const range = sheet.getDataRange();
      const formulas = range.getFormulas();
      const rich = range.getRichTextValues();
      const sid = sheet.getSheetId();

      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          const f = formulas[r][c];
          if (f) {
            const nf = ctx.remap(f);
            if (nf !== f) sheet.getRange(r + 1, c + 1).setFormula(nf);
            continue;
          }
          if (chipCells.has(`${sid}:${r}:${c}`)) continue;
          const rtv = rich[r][c];
          if (!rtv || !rtv.getText()) continue;
          const needs = ctx.idTest.test(rtv.getText()) ||
            rtv.getRuns().some(run => { const u = run.getLinkUrl(); return !!u && ctx.idTest.test(u); });
          if (needs) sheet.getRange(r + 1, c + 1).setRichTextValue(remapRichText(rtv, ctx));
        }
      }
    });
  }

  ss.createTextFinder(ctx.pattern)
    .useRegularExpression(true)
    .matchCase(CONFIG.MATCH_CASE)
    .matchFormulaText(true)   // false would overwrite formulas with their displayed value
    .replaceAllWith(ctx.newTerm);

  ss.getSheets().forEach(sh => {
    const n = sh.getName(), nn = ctx.replaceKw(n);
    if (nn !== n) sh.setName(nn);
  });

  SpreadsheetApp.flush();
}

function remapRichText(rtv, ctx) {
  let text = "";
  const spans = [];
  rtv.getRuns().forEach(run => {
    const t = ctx.remap(run.getText());
    const s = text.length;
    text += t;
    const link = run.getLinkUrl();
    spans.push({ s, e: text.length, style: run.getTextStyle(), link: link ? ctx.remap(link) : null });
  });
  const b = SpreadsheetApp.newRichTextValue().setText(text);
  spans.forEach(sp => {
    if (sp.e <= sp.s) return;
    b.setTextStyle(sp.s, sp.e, sp.style);
    if (sp.link) b.setLinkUrl(sp.s, sp.e, sp.link);
  });
  return b.build();
}

function rewriteChips(ssId, ctx) {
  const chipCells = new Set();
  if (typeof Sheets === "undefined") {
    ctx.warn("Advanced Sheets service not enabled, so smart chips were not rewritten.");
    return chipCells;
  }
  const resp = Sheets.Spreadsheets.get(ssId, {
    includeGridData: true,
    fields: "sheets(properties(sheetId,title),data(startRow,startColumn,rowData(values(userEnteredValue,chipRuns))))",
  });

  const requests = [];
  (resp.sheets || []).forEach(sh => {
    const sheetId = sh.properties.sheetId;
    (sh.data || []).forEach(block => {
      const r0 = block.startRow || 0, c0 = block.startColumn || 0;
      (block.rowData || []).forEach((row, ri) => {
        (row.values || []).forEach((cell, ci) => {
          if (!cell.chipRuns || !cell.chipRuns.some(run => run.chip)) return;
          const r = r0 + ri, c = c0 + ci;
          chipCells.add(`${sheetId}:${r}:${c}`);
          const rebuilt = rebuildChipCell(cell, ctx, `${sh.properties.title}!R${r + 1}C${c + 1}`);
          if (!rebuilt) return;
          requests.push({
            updateCells: {
              range: { sheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: c, endColumnIndex: c + 1 },
              rows: [{ values: [rebuilt] }],
              fields: "userEnteredValue,chipRuns",
            },
          });
        });
      });
    });
  });

  for (let i = 0; i < requests.length; i += 200) {
    Sheets.Spreadsheets.batchUpdate({ requests: requests.slice(i, i + 200) }, ssId);
  }
  return chipCells;
}

function rebuildChipCell(cell, ctx, where) {
  const text = (cell.userEnteredValue && cell.userEnteredValue.stringValue) || "";
  const runs = cell.chipRuns.slice().sort((a, b) => (a.startIndex || 0) - (b.startIndex || 0));
  let out = "", changed = false, unwritable = false;
  const newRuns = [];

  const pushText = seg => {
    const ns = ctx.replaceKw(ctx.remap(seg));
    if (ns !== seg) changed = true;
    out += ns;
  };

  const first = runs[0].startIndex || 0;
  if (first > 0) pushText(text.substring(0, first));

  for (let i = 0; i < runs.length; i++) {
    const start = runs[i].startIndex || 0;
    const end = i + 1 < runs.length ? (runs[i + 1].startIndex || 0) : text.length;
    const chip = runs[i].chip;
    if (!chip) { pushText(text.substring(start, end)); continue; }

    let newChip;
    if (chip.richLinkProperties) {
      const uri = chip.richLinkProperties.uri;
      const nu = ctx.remap(uri);
      if (nu !== uri) changed = true;
      if (!/^https:\/\/(docs|drive)\.google\.com\//.test(nu)) unwritable = true;
      newChip = { richLinkProperties: { uri: nu } };
    } else if (chip.personProperties) {
      const p = chip.personProperties;
      newChip = { personProperties: { email: p.email } };
      if (p.displayFormat) newChip.personProperties.displayFormat = p.displayFormat;
    } else {
      unwritable = true;
      newChip = chip;
    }
    newRuns.push({ startIndex: out.length, chip: newChip });
    out += "@";
  }

  if (!changed) return null;
  if (unwritable) {
    ctx.warn(`${where}: contains a non-Drive smart chip, which the API can't write. Update this cell by hand.`);
    return null;
  }
  return { userEnteredValue: { stringValue: out }, chipRuns: newRuns };
}
