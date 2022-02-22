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
    if (lastRow === MembersSheet.NUM_HEADER_ROWS) {
      return null;
    }

    return this.sheet.getRange(MembersSheet.NUM_HEADER_ROWS + 1, MembersSheet.COLUMN_INDEX_NAME, lastRow - MembersSheet.NUM_HEADER_ROWS);
  }
}

class Members {
  membersRange: GoogleAppsScript.Spreadsheet.Range;  // List of members in range format

  constructor(membersSheet: MembersSheet) {
    this.load(membersSheet);
  }

  load(membersSheet: MembersSheet) {
    this.membersRange = membersSheet.getMembersRange();
  }

  getMembersRange(): GoogleAppsScript.Spreadsheet.Range {
    return this.membersRange;
  }

  getNumMembers(): number {
    return this.membersRange ? this.membersRange.getNumRows() : 0;
  }

  getMembers(): string[] {
    return this.membersRange ? this.getMembersRange().getValues().flat() : [];
  }

  getMembersRef(): string[] {
    if (!this.membersRange) {
      return [];
    }
    return [...Array(this.getNumMembers())].map((_, i) => `R${MembersSheet.NUM_HEADER_ROWS + i + 1}C${MembersSheet.COLUMN_INDEX_NAME}`);
  }
}