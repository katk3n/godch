class PaymentsSheet {
  static readonly NUM_HEADER_ROWS = 1;
  static readonly NUM_FOOTER_ROWS = 1;
  static readonly ROW_INDEX_FROM = 1;

  sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payments");
  }

  private createTable(membersSheet: MembersSheet, detailsSheet: DetailsSheet) {
    const membersRange = membersSheet.getMembersRange();
    const numMembers = membersRange.getNumRows();
    for (let fromIndex = 1; fromIndex <= numMembers; fromIndex++) {
      // Create header
      const columnIndex = fromIndex * 2 - 1;  // member name is set to 1st, 3rd, 5th, ... column

      const refFromMemberName = `Members!${membersRange.getCell(fromIndex, 1).getA1Notation()}`;
      this.sheet.getRange(PaymentsSheet.ROW_INDEX_FROM, columnIndex).setFormula(`${refFromMemberName}&"から"`);
      this.sheet.getRange(PaymentsSheet.ROW_INDEX_FROM, columnIndex, 1, 2).merge()
        .setHorizontalAlignment("center")
        .setBackground(Constant.BGCOLOR_MEMBER);

      for (let toIndex = 1; toIndex <= numMembers; toIndex++) {
        const rowIndex = PaymentsSheet.NUM_HEADER_ROWS + toIndex;
        const refToMemberName = `Members!${membersRange.getCell(toIndex, 1).getA1Notation()}`;
        if (toIndex === fromIndex) {
          this.sheet.getRange(rowIndex, columnIndex)
            .setValue("支払済")
            .setFontColor(Constant.FONTCOLOR_INACTIVE)
            .setBackground(Constant.BGCOLOR_INACTIVE);
        } else {
          this.sheet.getRange(rowIndex, columnIndex).setFormula(`${refToMemberName}&"に"`);
        }

        const buyerRangeStr = `Details!${detailsSheet.getBuyerRange().getA1Notation()}`;
        const paymentRangeStr = `Details!${detailsSheet.getPaymentRange(fromIndex).getA1Notation()}`;
        const paymentFormula = `SUMIF(${buyerRangeStr},${refToMemberName},${paymentRangeStr})`;
        this.sheet.getRange(rowIndex, columnIndex + 1)
          .setFormula(paymentFormula)
          .setNumberFormat(Constant.FORMAT_CURRENCY);
        
        if (toIndex === fromIndex) {
          this.sheet.getRange(rowIndex, columnIndex + 1)
            .setFontColor(Constant.FONTCOLOR_INACTIVE)
            .setBackground(Constant.BGCOLOR_INACTIVE);
        }
      }

      /* Create cell for sum */
      const sumStartCellStr = `R${PaymentsSheet.NUM_HEADER_ROWS + 1}C${columnIndex + 1}`;
      const sumEndCellStr = `R${PaymentsSheet.NUM_HEADER_ROWS + numMembers}C${columnIndex + 1}`;
      const sumFormula = `SUM(${sumStartCellStr}:${sumEndCellStr})`;
      this.sheet.getRange(PaymentsSheet.NUM_HEADER_ROWS + numMembers + 1, columnIndex)
        .setValue("合計")
        .setBackground(Constant.BGCOLOR_SUM);
      this.sheet.getRange(PaymentsSheet.NUM_HEADER_ROWS + numMembers + 1, columnIndex + 1)
        .setFormulaR1C1(sumFormula)
        .setNumberFormat(Constant.FORMAT_CURRENCY)
        .setBackground(Constant.BGCOLOR_SUM);
    }
  }

  initialize() {
    const membersSheet = new MembersSheet();
    if (membersSheet.getMembers().length < 2) {
      throw new Error("2名以上の参加者を入力してください");
    }
    const detailsSheet = new DetailsSheet();
    this.sheet.clear();
    this.createTable(membersSheet, detailsSheet);
  }

  clear() {
    this.sheet.clear();
  }
}