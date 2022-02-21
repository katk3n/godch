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
  static readonly TITLE_RATIO = "費用負担率";
  static readonly TITLE_PAYMENT = "負担額";
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
      this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, columnIndex).setValue(DetailsSheet.TITLE_RATIO);
      this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, columnIndex + 1).setValue(DetailsSheet.TITLE_PAYMENT);

      /* Create cells for ratio */
      for (let j = 1; j <= DetailsSheet.NUM_INPUT_ROWS; j++) {
        const rowIndex = DetailsSheet.NUM_HEADER_ROWS + j;
        const priceCellStr = `R${rowIndex}C${DetailsSheet.COLUMN_INDEX_PRICE}`;
        const ratioCellStr = `R${rowIndex}C${columnIndex}`;
        const formula = `ROUND(${priceCellStr}*${ratioCellStr})`;
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

  private createCellsForColumnSum(membersSheet: MembersSheet) {
    const membersRange = membersSheet.getMembersRange();
    const numMembers = membersRange.getNumRows();
    const columnIndexRatio = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2) + 1;
    const columnIndexPayment = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2) + 2;

    /* Create header */
    this.sheet.getRange(DetailsSheet.ROW_INDEX_NAME, columnIndexRatio, 1, 2)
      .merge()
      .setValue("合計")
      .setHorizontalAlignment("center")
      .setBackground(Constant.BGCOLOR_SUM);

    this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, columnIndexRatio)
      .setValue(DetailsSheet.TITLE_RATIO)
      .setBackground(Constant.BGCOLOR_SUM);

    this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, columnIndexPayment)
      .setValue(DetailsSheet.TITLE_PAYMENT)
      .setBackground(Constant.BGCOLOR_SUM);

    /* Create cells for sum of each row */
    for (let i = 1; i <= DetailsSheet.NUM_INPUT_ROWS; i++) {
      const rowIndex = DetailsSheet.NUM_HEADER_ROWS + i;
      const headStartCellStr = `R${DetailsSheet.ROW_INDEX_HEAD}C${DetailsSheet.NUM_LEFT_FIXED_COLUMNS + 1}`;
      const headEndCellStr = `R${DetailsSheet.ROW_INDEX_HEAD}C${DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2)}`;
      const sumStartCellStr = `R${rowIndex}C${DetailsSheet.NUM_LEFT_FIXED_COLUMNS + 1}`;
      const sumEndCellStr = `R${rowIndex}C${DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2)}`;

      const ratioFormula = `SUMIF(${headStartCellStr}:${headEndCellStr},"${DetailsSheet.TITLE_RATIO}",${sumStartCellStr}:${sumEndCellStr})`;
      const paymentFormula = `SUMIF(${headStartCellStr}:${headEndCellStr},"${DetailsSheet.TITLE_PAYMENT}",${sumStartCellStr}:${sumEndCellStr})`;

      this.sheet.getRange(rowIndex, columnIndexRatio)
        .setFormulaR1C1(ratioFormula)
        .setBackground(Constant.BGCOLOR_SUM);

      this.sheet.getRange(rowIndex, columnIndexPayment)
        .setFormulaR1C1(paymentFormula)
        .setNumberFormat(Constant.FORMAT_CURRENCY)
        .setBackground(Constant.BGCOLOR_SUM);

      this.setRatioConditionalFormat(rowIndex, columnIndexRatio);
      this.setPaymentConditionalFormat(rowIndex, columnIndexPayment);
    }

    /* Create cell for total sum */
    const sumStartCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + 1}C${columnIndexPayment}`;
    const sumEndCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS}C${columnIndexPayment}`;
    const sumFormula = `SUM(${sumStartCellStr}:${sumEndCellStr})`;
    this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, columnIndexPayment)
      .setFormulaR1C1(sumFormula)
      .setBackground(Constant.BGCOLOR_SUM);
    this.setPaymentConditionalFormat(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, columnIndexPayment);

    // Empty cell, but just set color
    this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, columnIndexRatio)
      .setBackground(Constant.BGCOLOR_SUM);
  }

  private setRatioConditionalFormat(rowIndex: number, columnIndex: number) {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberNotEqualTo(1)
      .setFontColor("red")
      .setRanges([this.sheet.getRange(rowIndex, columnIndex)])
      .build();
    
    const rules = this.sheet.getConditionalFormatRules()
    rules.push(rule);
    this.sheet.setConditionalFormatRules(rules);
  }

  private setPaymentConditionalFormat(rowIndex: number, columnIndex: number) {
    const refPrice = this.sheet.getRange(rowIndex, DetailsSheet.COLUMN_INDEX_PRICE).getA1Notation();
    const refSum = this.sheet.getRange(rowIndex, columnIndex).getA1Notation();
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=NE(${refSum},${refPrice})`)
      .setFontColor("red")
      .setRanges([this.sheet.getRange(rowIndex, columnIndex)])
      .build();
    
    const rules = this.sheet.getConditionalFormatRules()
    rules.push(rule);
    this.sheet.setConditionalFormatRules(rules);
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
    this.createCellsForColumnSum(membersSheet);
  }

  clear() {
    this.sheet.clear();
  }
}