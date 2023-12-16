import { DetailsSheet } from "./details";
import { Util } from "./util";

export class PaymentsSheet {
  static readonly NUM_HEADER_ROWS = 1;
  static readonly NUM_FOOTER_ROWS = 1;
  static readonly ROW_INDEX_FROM = 1;

  sheet: GoogleAppsScript.Spreadsheet.Sheet | null;

  constructor() {
    this.sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payments");
  }

  /**
   * createHeader()
   * @param members
   */
  private createHeader(members: Members) {
    if (this.sheet == null) {
      console.error("PaymentsSheet is null");
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
      .getRange(PaymentsSheet.ROW_INDEX_FROM, 1, 1, numMembers * 2)
      .setValues(headerTable);

    for (let i = 0; i < numMembers; i++) {
      const columnIndex = i * 2 + 1; // member name is set to 1st, 3rd, 5th, ... column
      this.sheet
        .getRange(PaymentsSheet.ROW_INDEX_FROM, columnIndex, 1, 2)
        .merge()
        .setHorizontalAlignment("center")
        .setBackground(Util.BGCOLOR_MEMBER);
    }
  }

  /**
   * createTable()
   * @param members
   * @param detailsSheet
   */
  private createTable(members: Members, detailsSheet: DetailsSheet) {
    if (this.sheet == null) {
      console.error("PaymentsSheet is null");
      return;
    }

    const numMembers = members.getNumMembers();

    // Prepare values
    const paymentTable: string[][] = Util.createTable(
      numMembers,
      numMembers * 2
    );
    const membersRef = members.getMembersRef();

    const buyerRange = detailsSheet.getBuyerRange();
    if (buyerRange == null) {
      console.error("buyerRange is null");
      return;
    }

    const refBuyerRange = `Details!${buyerRange.getA1Notation()}`;

    for (let fromIndex = 0; fromIndex < numMembers; fromIndex++) {
      const paymentRange = detailsSheet.getPaymentRange(fromIndex + 1);
      if (paymentRange == null) {
        console.error(`paymentRange(${fromIndex + 1}) is null`);
        return;
      }

      const refPaymentRange = `Details!${paymentRange.getA1Notation()}`;

      for (let toIndex = 0; toIndex < numMembers; toIndex++) {
        const refToMemberName = `Members!${membersRef[toIndex]}`;
        paymentTable[toIndex][fromIndex * 2] =
          fromIndex === toIndex ? "支払済" : `=${refToMemberName}&"に"`;

        const paymentFormula = `=SUMIF(${refBuyerRange},${refToMemberName},${refPaymentRange})`;
        paymentTable[toIndex][fromIndex * 2 + 1] = paymentFormula;
      }
    }

    // Write values to spreadsheet
    this.sheet
      .getRange(
        PaymentsSheet.NUM_HEADER_ROWS + 1,
        1,
        numMembers,
        numMembers * 2
      )
      .setValues(paymentTable)
      .setNumberFormat(Util.FORMAT_CURRENCY);

    // Set color and format
    for (let fromIndex = 0; fromIndex < numMembers; fromIndex++) {
      const columnIndex = fromIndex * 2 + 1; // member name is set to 1st, 3rd, 5th, ... column

      this.sheet
        .getRange(
          PaymentsSheet.NUM_HEADER_ROWS + fromIndex + 1,
          columnIndex,
          1,
          2
        )
        .setFontColor(Util.FONTCOLOR_INACTIVE)
        .setBackground(Util.BGCOLOR_INACTIVE);

      this.sheet
        .getRange(
          PaymentsSheet.NUM_HEADER_ROWS + 1,
          columnIndex + 1,
          numMembers,
          1
        )
        .setNumberFormat(Util.FORMAT_CURRENCY);
    }
  }

  /**
   * createCellsForTotalPayment
   * @param members
   */
  createCellsForTotalPayment(members: Members) {
    if (this.sheet == null) {
      console.error("PaymentsSheet is null");
      return;
    }

    const numMembers = members.getNumMembers();

    // Prepare values
    const sumTable: string[][] = Util.createTable(1, numMembers * 2);

    for (let i = 0; i < numMembers; i++) {
      const columnIndex = i * 2 + 1; // member names are set to 1st, 3rd, 5th, ... columns of spreadsheet
      const sumStartCellStr = `R${PaymentsSheet.NUM_HEADER_ROWS + 1}C${
        columnIndex + 1
      }`;
      const sumEndCellStr = `R${PaymentsSheet.NUM_HEADER_ROWS + numMembers}C${
        columnIndex + 1
      }`;
      const sumFormula = `=SUM(${sumStartCellStr}:${sumEndCellStr})`;
      sumTable[0][i * 2] = "合計";
      sumTable[0][i * 2 + 1] = sumFormula;
    }

    // Write values to spreadsheet
    this.sheet
      .getRange(
        PaymentsSheet.NUM_HEADER_ROWS + numMembers + 1,
        1,
        1,
        numMembers * 2
      )
      .setValues(sumTable)
      .setBackground(Util.BGCOLOR_SUM);

    for (let i = 0; i < numMembers; i++) {
      const columnIndex = i * 2 + 1; // member names are set to 1st, 3rd, 5th, ... columns of spreadsheet
      this.sheet
        .getRange(
          PaymentsSheet.NUM_HEADER_ROWS + numMembers + 1,
          columnIndex + 1
        )
        .setNumberFormat(Util.FORMAT_CURRENCY);
    }
  }

  /**
   * initialize()
   */
  initialize() {
    if (this.sheet == null) {
      console.error("PaymentsSheet is null");
      return;
    }

    const membersSheet = new MembersSheet();
    const members = new Members(membersSheet);
    if (members.getNumMembers() < 2) {
      throw new Error("2名以上の参加者を入力してください");
    }
    const detailsSheet = new DetailsSheet();
    this.sheet.clear();
    this.createHeader(members);
    this.createTable(members, detailsSheet);
    this.createCellsForTotalPayment(members);
  }

  /**
   * clear()
   */
  clear() {
    if (this.sheet == null) {
      console.error("PaymentsSheet is null");
      return;
    }

    this.sheet.clear();
  }
}
