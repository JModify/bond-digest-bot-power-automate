// Reserve a group post before sending and record completion after Outlook accepts the draft.
// One flow with sequential runs owns this worksheet.

const worksheetName = "Digest Bot Group Posts";
const headers = ["Post key", "State", "Updated UTC"];
const pending = "Pending";
const processed = "Processed";

function main(workbook: ExcelScript.Workbook, postId: string, recordProcessed: boolean): boolean {
  if (!postId.trim()) throw new Error("The group post ID is required.");

  // A fixed prefix makes every ID an Excel text value, including IDs that start with a symbol.
  const key = "post:" + postId;
  let sheet = workbook.getWorksheet(worksheetName);
  if (!sheet) {
    if (recordProcessed) throw new Error("The group post was never claimed.");
    sheet = workbook.addWorksheet(worksheetName);
  }

  const used = sheet.getUsedRange(true);
  const updatedAt = new Date().toISOString();
  if (!used) {
    if (recordProcessed) throw new Error("The group post was never claimed.");
    sheet.getRange("A1:C2").setValues([headers, [key, pending, updatedAt]]);
    return true;
  }

  const rows = used.getValues();
  if (
    used.getRowIndex() !== 0 ||
    used.getColumnIndex() !== 0 ||
    used.getColumnCount() !== headers.length ||
    !headers.every((header, column) => rows[0][column] === header)
  ) {
    throw new Error("Check the headers on the Digest Bot Group Posts worksheet.");
  }

  const index = rows.findIndex((row) => row[0] === key);
  if (recordProcessed) {
    if (index === -1) throw new Error("The group post was never claimed.");
    if (rows[index][1] !== processed) {
      sheet.getRangeByIndexes(index, 1, 1, 2).setValues([[processed, updatedAt]]);
    }
    return true;
  }

  // Both an accepted send and a pending send block automatic repetition of the same post.
  if (index !== -1) return false;
  sheet.getRangeByIndexes(rows.length, 0, 1, headers.length).setValues([[key, pending, updatedAt]]);
  return true;
}
