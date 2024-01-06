import { Members, MembersSheet } from "./members";
import { Util } from "./util";

export class BalanceSheet {
  static readonly NUM_HEADER_ROWS = 1;
  static readonly ROW_INDEX_FROM = 1;

  sheet: GoogleAppsScript.Spreadsheet.Sheet | null;

  constructor() {
    this.sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Balance");
  }

  /**
   * createHeader()
   * @param members
   */
  private createHeader(members: Members) {
    if (this.sheet == null) {
      console.error("BalanceSheet is null");
      return;
    }

    const numMembers = members.getNumMembers();

    // Prepare values
    // member names are set to even columns of table, odd columns are empty
    const headerTable: string[][] = Util.createTable(1, numMembers * 2);
    const membersRef = members.getMembersRef();
    for (let i = 0; i < numMembers; i++) {
      const columnIndex = i * 2; // member names are set to 0th, 2nd, 4th, ... column of table
      const refFromMemberName = `Members!${membersRef[i]}`;
      headerTable[0][columnIndex] = `=${refFromMemberName}&"から"`;
    }

    // Write values to spreadsheet
    this.sheet
      .getRange(BalanceSheet.ROW_INDEX_FROM, 1, 1, numMembers * 2)
      .setValues(headerTable);

    for (let i = 0; i < numMembers; i++) {
      const columnIndex = i * 2 + 1; // member name is set to 1st, 3rd, 5th, ... column
      this.sheet
        .getRange(BalanceSheet.ROW_INDEX_FROM, columnIndex, 1, 2)
        .merge()
        .setHorizontalAlignment("center")
        .setBackground(Util.BGCOLOR_MEMBER);
    }
  }

  /**
   * createTable()
   * @param members
   * @param balanceSheet
   */
  private createTable(members: Members, balanceSheet: BalanceSheet) {
    if (this.sheet == null) {
      console.error("BalanceSheet is null");
      return;
    }

    const numMembers = members.getNumMembers();

    // Prepare values
    const balanceTable: string[][] = Util.createTable(
      numMembers,
      numMembers * 2
    );
    const membersRef = members.getMembersRef();

    for (let fromIndex = 0; fromIndex < numMembers; fromIndex++) {
      for (let toIndex = 0; toIndex < numMembers; toIndex++) {
        const refToMemberName = `Members!${membersRef[toIndex]}`;
        balanceTable[toIndex][fromIndex * 2] =
          fromIndex === toIndex ? "" : `=${refToMemberName}&"に"`;

        // Index of spreadsheets starts from 1
        const toPay = `Payments!R${
          toIndex + 1 + BalanceSheet.NUM_HEADER_ROWS
        }C${(fromIndex + 1) * 2}`;
        const toBePaid = `Payments!R${
          fromIndex + 1 + BalanceSheet.NUM_HEADER_ROWS
        }C${(toIndex + 1) * 2}`;

        const balanceFormula = `=MAX(${toPay}-${toBePaid},0)`;
        balanceTable[toIndex][fromIndex * 2 + 1] = balanceFormula;
      }
    }

    // Write values to spreadsheet
    this.sheet
      .getRange(BalanceSheet.NUM_HEADER_ROWS + 1, 1, numMembers, numMembers * 2)
      .setValues(balanceTable)
      .setNumberFormat(Util.FORMAT_CURRENCY);

    // Set color and format
    for (let fromIndex = 0; fromIndex < numMembers; fromIndex++) {
      const columnIndex = fromIndex * 2 + 1; // member name is set to 1st, 3rd, 5th, ... column

      this.sheet
        .getRange(
          BalanceSheet.NUM_HEADER_ROWS + fromIndex + 1,
          columnIndex,
          1,
          2
        )
        .setFontColor(Util.FONTCOLOR_INACTIVE)
        .setBackground(Util.BGCOLOR_INACTIVE);

      this.sheet
        .getRange(
          BalanceSheet.NUM_HEADER_ROWS + 1,
          columnIndex + 1,
          numMembers,
          1
        )
        .setNumberFormat(Util.FORMAT_CURRENCY);
    }
  }

  /**
   * initialize()
   */
  initialize() {
    if (this.sheet == null) {
      console.error("BalanceSheet is null");
      return;
    }

    const membersSheet = new MembersSheet();
    const members = new Members(membersSheet);
    if (members.getNumMembers() < 2) {
      throw new Error("2名以上の参加者を入力してください");
    }
    const balanceSheet = new BalanceSheet();
    this.sheet.clear();
    this.createHeader(members);
    this.createTable(members, balanceSheet);
  }

  /**
   * clear()
   */
  clear() {
    if (this.sheet == null) {
      console.error("BalanceSheet is null");
      return;
    }

    this.sheet.clear();
  }
}
