class DetailsSheet {
  static readonly NUM_HEADER_ROWS = 2;
  static readonly NUM_INPUT_ROWS = 30;
  static readonly NUM_FOOTER_ROWS = 1;
  static readonly NUM_LEFT_FIXED_COLUMNS = 4;
  static readonly COLUMN_INDEX_DATE = 1;
  static readonly COLUMN_INDEX_BUYER = 2;
  static readonly COLUMN_INDEX_DETAIL = 3;
  static readonly COLUMN_INDEX_PRICE = 4;
  static readonly ROW_INDEX_NAME = 1;
  static readonly ROW_INDEX_HEAD = 2;
  sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Details");
  }

  private createLeftHeader() {
    this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, DetailsSheet.COLUMN_INDEX_DATE).setValue("発生日");
    this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, DetailsSheet.COLUMN_INDEX_BUYER).setValue("立替者");
    this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, DetailsSheet.COLUMN_INDEX_DETAIL).setValue("内容");
    this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, DetailsSheet.COLUMN_INDEX_PRICE).setValue("金額");
  }

  private createDropDown(membersSheet: MembersSheet) {
    const range = this.sheet.getRange(
      DetailsSheet.NUM_HEADER_ROWS + 1, DetailsSheet.COLUMN_INDEX_BUYER,
      DetailsSheet.NUM_INPUT_ROWS, 1);
    range.clearDataValidations();
    const rule = SpreadsheetApp.newDataValidation();
    rule.requireValueInRange(membersSheet.getMembersRange());
    range.setDataValidation(rule);
  }

  private createRatioInputForm(membersSheet: MembersSheet) {
    const membersRange = membersSheet.getMembersRange();
    const numMembers = membersRange.getNumRows();

    for (let i = 1; i <= numMembers; i++) {
      // member name is set to 5th, 7th, 9th, ... column
      const columnIndex = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + i * 2 - 1;

      /* Create header */
      // Refer to the member name from Members sheet
      this.sheet.getRange(DetailsSheet.ROW_INDEX_NAME, columnIndex).setFormula(`Members!${membersRange.getCell(i, 1).getA1Notation()}`);
      this.sheet.getRange(DetailsSheet.ROW_INDEX_NAME, columnIndex, 1, 2).merge()
        .setHorizontalAlignment("center")
        .setBackground(Constant.BGCOLOR_MEMBER);
      this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, columnIndex).setValue("費用負担率");
      this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, columnIndex + 1).setValue("負担額");


      /* Create cells for ratio */
      for (let j = 1; j <= DetailsSheet.NUM_INPUT_ROWS; j++) {
        const rowIndex = DetailsSheet.NUM_HEADER_ROWS + j;
        const priceCellStr = `R${rowIndex}C${DetailsSheet.COLUMN_INDEX_PRICE}`;
        const ratioCellStr = `R${rowIndex}C${columnIndex}`;
        const formula = `${priceCellStr}*${ratioCellStr}`;
        this.sheet.getRange(rowIndex, columnIndex + 1)
          .setFormulaR1C1(formula)
          .setNumberFormat(Constant.FORMAT_CURRENCY)
          .setBackground(Constant.BGCOLOR_INACTIVE);
      }

      /* Create cell for sum */
      const sumStartCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + 1}C${columnIndex + 1}`;
      const sumEndCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS}C${columnIndex + 1}`;
      const sumFormula = `SUM(${sumStartCellStr}:${sumEndCellStr})`;
      this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, columnIndex + 1)
        .setFormulaR1C1(sumFormula)
        .setBackground(Constant.BGCOLOR_SUM);

      // Empty cell, but just set color
      this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, columnIndex)
        .setBackground(Constant.BGCOLOR_SUM);
    }

    /* Create cell for total sum */
      const sumStartCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + 1}C${DetailsSheet.COLUMN_INDEX_PRICE}`;
      const sumEndCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS}C${DetailsSheet.COLUMN_INDEX_PRICE}`;
      const sumFormula = `SUM(${sumStartCellStr}:${sumEndCellStr})`;
      this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, DetailsSheet.COLUMN_INDEX_PRICE - 1)
        .setValue("合計")
        .setBackground(Constant.BGCOLOR_SUM);
      this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, DetailsSheet.COLUMN_INDEX_PRICE)
        .setFormulaR1C1(sumFormula)
        .setBackground(Constant.BGCOLOR_SUM);
  }

  getBuyerRange(): GoogleAppsScript.Spreadsheet.Range {
    return this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + 1, DetailsSheet.COLUMN_INDEX_BUYER, DetailsSheet.NUM_INPUT_ROWS, 1);
  }

  getPaymentRange(memberIndex: number): GoogleAppsScript.Spreadsheet.Range {
    const columnIndex = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (memberIndex * 2)
    return this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + 1, columnIndex, DetailsSheet.NUM_INPUT_ROWS, 1);
  }

  initialize() {
    const membersSheet = new MembersSheet();
    if (membersSheet.getMembers().length < 2) {
      throw new Error("2名以上の参加者を入力してください");
    }
    this.sheet.clear();
    this.createLeftHeader();
    this.createDropDown(membersSheet);
    this.createRatioInputForm(membersSheet);
  }

  clear() {
    this.sheet.clear();
  }
}