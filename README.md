# UPE Induction Suite Script

An Apps Script that sets up the Google Drive folder for a new induction quarters of [Upsilon Pi Epsilon at UCLA](https://upe.seas.ucla.edu/). It copies the template folder, renames everything for the new term, locks forms to org accounts, and fixes the `IMPORTRANGE` links so the new sheets point at each other instead of the template. It also works as a general tool for deep-copying a Drive folder.

## What it does

1. Copies the template folder and everything in it.
2. Replaces the placeholder keyword (`TERM` by default) in file names and formulas with the new quarter name, e.g. `Fall 2026`.
3. Turns on org-only login and email collection for every copied form.
4. Rewrites formulas and `IMPORTRANGE` references to point at the new copies.

The template itself is never modified. The script only calls `makeCopy()` on it and never writes to the original files. To be safe, share the template folder with the team as **Viewer** so Drive will block any accidental writes.

## Running it

1. Open the I&M Suite Script.
2. Go to **Induction Setup → Provision New Quarter...**
3. Enter the template folder ID (the part after `/folders/` in the folder's URL), or keep the default.
4. Enter the new quarter name and click **Confirm**.

When it finishes, you'll get a link to the new folder.

The first time you run it, Google will ask for permission to manage your files. Click **Review Permissions**, pick your org account, and click **Allow**.

## Development

The code lives on GitHub and is synced to Apps Script with the [Google Apps Script GitHub Assistant](https://chromewebstore.google.com/detail/google-apps-script-github/lfjcgcmkmjjlieihflfhjopckgpelofo) Chrome extension.

**Setup:** Install the extension, create a GitHub personal access token (classic) with `repo` scope, then open **Extensions → Apps Script** from the sheet. Paste the token into the extension settings and link it to `your-org/induction-quarter-provisioner` on `main` (or your branch).

**Pulling:** Select the branch in the GitHub toolbar above the editor and click **Pull (↓)**.

**Pushing:** After testing your changes in the editor, enter a commit message in the toolbar and click **Push (↑)**.

## Configuration

Edit the `CONFIG` object at the top of `Code.gs`:

```javascript
const CONFIG = {
  DEFAULT_TEMPLATE_FOLDER_ID: "1aBcDeFgHiJkLmNoPqRsTuVwXyZ12345", // pre-fills the prompt
  DESTINATION_PARENT_FOLDER_ID: "",  // blank = create new folder next to the template
  KEYWORD_TO_REPLACE: "TERM",        // placeholder in file names and formulas
  REQUIRE_ORG_LOGIN: true,           // responders must sign in with an org account
  COLLECT_RESPONDER_EMAIL: true,     // record the signed-in email
  LIMIT_ONE_RESPONSE: false          // one submission per person
};
```
