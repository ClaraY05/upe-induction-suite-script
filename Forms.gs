function processForm(formId, ctx) {
  ctx.assertWritable(formId);
  const form = FormApp.openById(formId);

  if (CONFIG.PUBLISH_FORMS && typeof form.isPublished === "function" && !form.isPublished()) {
    form.setPublished(true);
  }
  if (CONFIG.REQUIRE_ORG_LOGIN) {
    try { form.setRequireLogin(true); } catch (e) { ctx.warn(`setRequireLogin failed (Workspace accounts only): ${e.message}`); }
  }
  if (CONFIG.COLLECT_RESPONDER_EMAIL) {
    if (typeof form.setEmailCollectionType === "function") {
      form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
    } else {
      form.setCollectEmail(true);
    }
  }
  if (CONFIG.LIMIT_ONE_RESPONSE) form.setLimitOneResponsePerUser(true);

  const fix = s => ctx.replaceKw(ctx.remap(s));
  const upd = (get, set) => { const v = get(); if (v) { const nv = fix(v); if (nv !== v) set(nv); } };

  upd(() => form.getTitle(), v => form.setTitle(v));
  upd(() => form.getDescription(), v => form.setDescription(v));
  upd(() => form.getConfirmationMessage(), v => form.setConfirmationMessage(v));

  const quiz = form.isQuiz();
  form.getItems().forEach(item => {
    upd(() => item.getTitle(), v => item.setTitle(v));
    upd(() => item.getHelpText(), v => item.setHelpText(v));
    if (!quiz) fixChoices(item, ctx);
  });

  try { return publishedFormKey(form.getPublishedUrl()); } catch (e) { return null; }
}

function fixChoices(item, ctx) {
  const T = FormApp.ItemType;
  const t = item.getType();
  let typed;
  if (t === T.MULTIPLE_CHOICE) typed = item.asMultipleChoiceItem();
  else if (t === T.CHECKBOX) typed = item.asCheckboxItem();
  else if (t === T.LIST) typed = item.asListItem();
  else return;

  const choices = typed.getChoices();
  if (!choices.some(ch => ctx.hasKw(ch.getValue()))) return;

  typed.setChoices(choices.map(ch => {
    const v = ctx.replaceKw(ch.getValue());
    if (t === T.CHECKBOX) return typed.createChoice(v);
    try {
      const page = ch.getGotoPage();
      if (page) return typed.createChoice(v, page);
      const nav = ch.getPageNavigationType();
      if (nav) return typed.createChoice(v, nav);
    } catch (e) {}
    return typed.createChoice(v);
  }));
}