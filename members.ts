class MembersSheet {
  static readonly COLUMN_INDEX_NAME = 2;
  static readonly NUM_HEADER_ROWS = 1;
  sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Members");
  }

  getMembersRange(): GoogleAppsScript.Spreadsheet.Range {
    const lastRow = this.sheet.getRange(this.sheet.getMaxRows(), MembersSheet.COLUMN_INDEX_NAME)
      .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    return this.sheet.getRange(MembersSheet.NUM_HEADER_ROWS + 1, MembersSheet.COLUMN_INDEX_NAME, lastRow - MembersSheet.NUM_HEADER_ROWS);
  }

  getMembers(): string[] {
    if (this.isEmpty()) {
      return [];
    }
    return this.getMembersRange().getValues().flat();
  }

  isEmpty(): boolean {
    const lastRow = this.sheet.getRange(this.sheet.getMaxRows(), MembersSheet.COLUMN_INDEX_NAME)
      .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    return lastRow === MembersSheet.NUM_HEADER_ROWS;
  }
}