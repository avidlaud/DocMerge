/**
 * Add DocMerge to toolbar
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('DocMerge')
    .addItem('Fill Doc Templates', 'fillTemplate')
    .addToUi();
}

function fillTemplate(sheet=SpreadsheetApp.getActiveSheet()) {
  const ui = SpreadsheetApp.getUi();

  const templatePrompt = ui.prompt("Input the name of the template document", Browser.Buttons.OK_CANCEL);
  const templateId = templatePrompt.getResponseText();
  // No template document defined
  if (templateId === "cancel" || templateId == "") {
    return;
  }

  const startPrompt = ui.prompt("Input the starting row", Browser.Buttons.OK_CANCEL);
  const startResp = startPrompt.getResponseText();
  // Starting row
  if (startResp === "cancel" || startResp == "") {
    return;
  }
  const startRow = parseInt(startResp);

  const endPrompt = ui.prompt("Input the ending row (inclusive)", Browser.Buttons.OK_CANCEL);
  const endResp = endPrompt.getResponseText();
  // Ending row
  if (endResp === "cancel" || endResp == "") {
    return;
  }
  const endRow = parseInt(endResp);

  const templateDoc = DocumentApp.openById(templateId);
  const templateKeys = getKeysFromTemplate(templateDoc);

  const data = sheet.getDataRange().getDisplayValues();
  // First row contains keys
  const keys = data[0];

  invalidKey = validateKeys(templateKeys, keys);
  if (invalidKey != null) {
    Logger.log("Invalid key!");
    Logger.log(invalidKey);
    return;
  }

  entries = [];
  // Iterate over the selected rows
  for (let i = startRow - 1; i < endRow; i++) {
    entries.push(createEntryMap(keys, data[i]));
  }

  for (const entry of entries) {
    const newDocId = copyDocument(templateId);
    const newDoc = DocumentApp.openById(newDocId);
    const newDocBody = newDoc.getBody();
    // Replace the text
    for (const k in entry) {
      if (entry.hasOwnProperty(k)) {
        // Match pattern
        const pattern = "\\$\\{\\@" + regexString(k) + "\\@\\}";
        newDocBody.replaceText(pattern, (entry[k] == "") ? "?":entry[k]);
      }
    }
  }

  function getKeysFromTemplate(tDoc) {
    const tDocBody = tDoc.getBody();
    const keys = [];
    let match;
    // Find all of the keys in the template
    while ((match = tDocBody.findText("\\$\\{\\@[^\\@]+\\@\\}", match)) != null) {
      const startIdx = match.getStartOffset();
      const endIdx = match.getEndOffsetInclusive();
      const text = match.getElement().asText().getText().substring(startIdx + 3, endIdx - 1);
      keys.push(text);
    }
    return keys;
  }

  /* Check that all of the keys in the template exist in the spreadsheet */
  function validateKeys(templateKeys, availableKeys) {
    for (let a of templateKeys) {
      if (!availableKeys.includes(a)) {
        return a
      }
    }
    return null
  }

  /* Create map given arrays of keys and values */
  function createEntryMap(keys, values) {
    const m = {};
    for (const [i, e] of keys.entries()) {
      m[e] = values[i];
    }
    return m;
  }

  /* Escape all special characters in the template key */
  function regexString(str) {
    return str.replaceAll(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function copyDocument(templateId) {
    let file = DriveApp.getFileById(templateId);
    const time = new Date().getTime();
    const copy = file.makeCopy(time.toString());
    return copy.getId();
  }
}
