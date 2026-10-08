// Rail Test Tracker -> Daily Report. Office Script for Excel (Automate tab > New script).
// It is run by a Power Automate flow with one text parameter, "payload", that holds the
// numbers sent by the app. It finds your label cells in column B, clears the old data
// to their right (columns C and D) and writes the new data, adding rows if needed.

// ---- settings you can change if your sheet changes ----
const LABEL_COLUMN = 1;   // column B (A=0, B=1, C=2 ...) holds the labels
const DATA_COLUMN = 2;    // data starts in column C and uses 2 columns (C and D)
const MAX_BLOCK_ROWS = 80; // safety stop: never clear more rows than this
// Text that identifies each label cell (case does not matter; the colon is ignored).
const TOTAL_LABEL = "total miles left";
const REMAINING_LABEL = "range of miles left";
const TESTED_LABEL = "range of miles tested";

function main(workbook: ExcelScript.Workbook, payload: string): string {
  let data: { kind: string; summary: string[][]; body: string[][]; bodyTitle: number[] };
  try {
    data = JSON.parse(payload);
  } catch (e) {
    return "ERROR: could not read the data sent by the app.";
  }
  if (!data || (data.kind !== "remaining" && data.kind !== "tested")) {
    return "ERROR: unknown kind of data.";
  }

  const messages: string[] = [];
  if (data.kind === "remaining") {
    const r1 = writeBlock(workbook, TOTAL_LABEL, REMAINING_LABEL, data.summary || [], [], true);
    messages.push(r1);
    const r2 = writeBlock(workbook, REMAINING_LABEL, null, data.body || [], data.bodyTitle || [], false);
    messages.push(r2);
  } else {
    const r = writeBlock(workbook, TESTED_LABEL, null, data.body || [], data.bodyTitle || [], false);
    messages.push(r);
  }
  return messages.join(" | ");
}

// Finds the row (0-based) of the first cell in the label column that contains the text.
function findLabelRow(sheet: ExcelScript.Worksheet, text: string): number {
  const used = sheet.getUsedRange();
  if (!used) return -1;
  const lastRow = used.getRowIndex() + used.getRowCount();
  const col = sheet.getRangeByIndexes(0, LABEL_COLUMN, lastRow, 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (norm(String(col[i][0])).indexOf(text) >= 0) return i;
  }
  return -1;
}

function norm(s: string): string {
  return s.toLowerCase().replace(/[:\s]+/g, " ").trim();
}

function lastUsedRow(sheet: ExcelScript.Worksheet): number {
  const used = sheet.getUsedRange();
  return used ? used.getRowIndex() + used.getRowCount() - 1 : 0;
}

// Fills the block that starts at the row of startLabel. The block runs until the row of
// endLabel (summary block) or until the next non-empty label in the label column.
function writeBlock(
  workbook: ExcelScript.Workbook,
  startLabel: string,
  endLabel: string | null,
  rows: string[][],
  titleRows: number[],
  mergeAll: boolean
): string {
  // find the sheet that has the label
  let sheet: ExcelScript.Worksheet | null = null;
  let start = -1;
  const sheets = workbook.getWorksheets();
  for (let s = 0; s < sheets.length; s++) {
    const row = findLabelRow(sheets[s], startLabel);
    if (row >= 0) { sheet = sheets[s]; start = row; break; }
  }
  if (!sheet || start < 0) return "ERROR: could not find a label containing '" + startLabel + "' in column B.";

  // where does the block end (exclusive)?
  const last = lastUsedRow(sheet);
  let end = -1;
  if (endLabel) {
    const e = findLabelRow(sheet, endLabel);
    if (e > start) end = e;
  }
  if (end < 0) {
    const col = sheet.getRangeByIndexes(0, LABEL_COLUMN, last + 1, 1).getValues();
    for (let i = start + 1; i < col.length; i++) {
      if (String(col[i][0]).trim() !== "") { end = i; break; }
    }
  }
  if (end < 0) end = last + 1;               // last block on the sheet
  let height = end - start;
  if (height > MAX_BLOCK_ROWS) return "ERROR: block under '" + startLabel + "' looks too tall (" + height + " rows); nothing changed.";

  // add rows if the data does not fit (new rows copy the look of the row above)
  const need = rows.length;
  if (need > height) {
    const extra = need - height;
    sheet.getRangeByIndexes(end, 0, extra, 1).getEntireRow().insert(ExcelScript.InsertShiftDirection.down);
    end += extra;
    height += extra;
  }
  if (need === 0 && height === 0) return "nothing to write";

  // clear the old data in the block
  const region = sheet.getRangeByIndexes(start, DATA_COLUMN, Math.max(height, 1), 2);
  region.unmerge();
  region.clear(ExcelScript.ClearApplyTo.contents);
  region.getFormat().getFont().setBold(false);
  region.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);

  if (need === 0) return "cleared " + startLabel;

  // write as text so Excel never turns things like 1-2 into dates
  const target = sheet.getRangeByIndexes(start, DATA_COLUMN, need, 2);
  target.setNumberFormat("@");
  const values: string[][] = rows.map(function (r) { return [r[0] == null ? "" : String(r[0]), r[1] == null ? "" : String(r[1])]; });
  target.setValues(values);

  if (mergeAll) {
    for (let i = 0; i < need; i++) sheet.getRangeByIndexes(start + i, DATA_COLUMN, 1, 2).merge(false);
  }
  for (let t = 0; t < titleRows.length; t++) {
    const idx = titleRows[t];
    if (idx >= 0 && idx < need) {
      const cell = sheet.getRangeByIndexes(start + idx, DATA_COLUMN, 1, 1).getFormat().getFont();
      cell.setBold(true);
      cell.setUnderline(ExcelScript.RangeUnderlineStyle.single);
    }
  }
  return "wrote " + need + " rows under '" + startLabel + "' (row " + (start + 1) + ")";
}
