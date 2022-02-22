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

  /**
   * createLeftHeader()
   */
  private createLeftHeader() {
    const headerValues = [["発生日", "立替者", "内容", "金額"]];
    this.sheet.getRange(DetailsSheet.ROW_INDEX_HEAD, 1, 1, DetailsSheet.NUM_LEFT_FIXED_COLUMNS).setValues(headerValues);
  }

  /**
   * createDropDown()
   * @param members
   */
  private createDropDown(members: Members) {
    const range = this.sheet.getRange(
      DetailsSheet.NUM_HEADER_ROWS + 1, DetailsSheet.COLUMN_INDEX_BUYER,
      DetailsSheet.NUM_INPUT_ROWS, 1);
    range.clearDataValidations();
    const rule = SpreadsheetApp.newDataValidation();
    rule.requireValueInRange(members.getMembersRange());
    range.setDataValidation(rule);
  }

  /**
   * createHeader()
   * @param members 
   */
  private createHeader(members: Members) {
    const numMembers = members.getNumMembers();

    // Prepare header values
    const headerTable: string[][] = Util.createTable(DetailsSheet.NUM_HEADER_ROWS, numMembers * 2);

    // array indexes start with 0 while spreadsheet indexes start with 1
    const rowIndexName = DetailsSheet.ROW_INDEX_NAME - 1;
    const rowIndexHead = DetailsSheet.ROW_INDEX_HEAD - 1;

    const membersRef = members.getMembersRef();

    for (let i = 0; i < numMembers; i++) {
      // member name is set to 0th, 2nd, 4th, ... column of table
      const columnIndex = i * 2;

      // Refer to the member name from Members sheet, indexes of spreadsheet start with 1
      headerTable[rowIndexName][columnIndex] = `=Members!${membersRef[i]}`;
      headerTable[rowIndexHead][columnIndex] = DetailsSheet.TITLE_RATIO;
      headerTable[rowIndexHead][columnIndex + 1] = DetailsSheet.TITLE_PAYMENT;
    }

    // Write header values to spreadsheet
    this.sheet.getRange(1, DetailsSheet.NUM_LEFT_FIXED_COLUMNS + 1, DetailsSheet.NUM_HEADER_ROWS, numMembers * 2).setValues(headerTable);

    // Merge each name cell with the next empty cell
    for (let i = 0; i < numMembers; i++) {
      // member name is set to 1st, 3rd, 5th, ... column of spreadsheet
      const columnIndex = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + i * 2 + 1;
      this.sheet.getRange(DetailsSheet.ROW_INDEX_NAME, columnIndex, 1, 2).merge()
        .setHorizontalAlignment("center")
        .setBackground(Util.BGCOLOR_MEMBER);
    }
  }

  /**
   * createRatioInputForm()
   * @param members 
   */
  private createRatioInputForm(members: Members) {
    const numMembers = members.getNumMembers();

    // Prepare values
    // even columns are for ratio, odd columns are for payment (= price * ratio)
    const inputTable: string[][] = Util.createTable(DetailsSheet.NUM_INPUT_ROWS, numMembers * 2);

    for (let i = 0; i < numMembers; i++) {
      // ratio is set to 5th, 7th, 9th, ... column of spreadsheet
      const columnIndex = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + i * 2 + 1;

      // Create cells for ratio
      for (let j = 0; j < DetailsSheet.NUM_INPUT_ROWS; j++) {
        const rowIndex = DetailsSheet.NUM_HEADER_ROWS + j + 1;
        const refPriceCell = `R${rowIndex}C${DetailsSheet.COLUMN_INDEX_PRICE}`;
        const refRatioCell = `R${rowIndex}C${columnIndex}`;
        const formula = `=ROUND(${refPriceCell}*${refRatioCell})`;

        // payment is set to 1st, 3rd, 5th, ... column of table
        inputTable[j][i * 2 + 1] = formula;
      }
    }

    // Write values to spreadsheet
    this.sheet.getRange(
      DetailsSheet.NUM_HEADER_ROWS + 1, DetailsSheet.NUM_LEFT_FIXED_COLUMNS + 1,
      DetailsSheet.NUM_INPUT_ROWS, numMembers * 2).setValues(inputTable);
    
    // Set format and background color
    for (let i = 0; i < numMembers; i++) {
      const columnIndex = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + i * 2 + 1;
      this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + 1, columnIndex + 1, DetailsSheet.NUM_INPUT_ROWS, 1)
        .setNumberFormat(Util.FORMAT_CURRENCY)
        .setBackground(Util.BGCOLOR_INACTIVE);
    }
  }

  /**
   * createCellsForTotalPrice()
   */
  private createCellsForTotalPrice() {
    const sumStartCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + 1}C${DetailsSheet.COLUMN_INDEX_PRICE}`;
    const sumEndCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS}C${DetailsSheet.COLUMN_INDEX_PRICE}`;
    const sumFormula = `=SUM(${sumStartCellStr}:${sumEndCellStr})`;
    const values = [["合計", sumFormula]];
    this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, DetailsSheet.COLUMN_INDEX_PRICE - 1, 1, 2)
      .setValues(values)
      .setBackground(Util.BGCOLOR_SUM);
  }

  /**
   * createCellsForTotalPaymentPerMember()
   * @param members 
   */
  private createCellsForTotalPaymentPerMember(members: Members) {
    const numMembers = members.getNumMembers();

    // Prepare values
    // even columns are empty, odd columns are for total payment
    const sumTable: string[][] = Util.createTable(1, numMembers * 2);

    for (let i = 0; i < numMembers; i++) {
      // each payment is set to 6th, 8th, 10th, ... column
      const columnIndex = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + i * 2 + 2;

      // Create cell for sum
      const sumStartCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + 1}C${columnIndex}`;
      const sumEndCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS}C${columnIndex}`;
      const sumFormula = `=SUM(${sumStartCellStr}:${sumEndCellStr})`;
      sumTable[0][2 * i + 1] = sumFormula;
    }

    // Write values to spreadsheet
    this.sheet.getRange(
      DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, DetailsSheet.NUM_LEFT_FIXED_COLUMNS + 1,
      1, numMembers * 2).setValues(sumTable).setBackground(Util.BGCOLOR_SUM);
  }

  /**
   * createHeaderForTotalPaymentPerItem()
   * @param members 
   */
  private createHeaderForTotalPaymentPerItem(members: Members) {
    const numMembers = members.getNumMembers();
    const columnIndexRatio = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2) + 1;

    const headerTable = [
      ["合計", ""],
      [DetailsSheet.TITLE_RATIO, DetailsSheet.TITLE_PAYMENT]
    ]

    this.sheet.getRange(DetailsSheet.ROW_INDEX_NAME, columnIndexRatio, 2, 2)
      .setValues(headerTable)
      .setBackground(Util.BGCOLOR_SUM);

    this.sheet.getRange(DetailsSheet.ROW_INDEX_NAME, columnIndexRatio, 1, 2)
      .merge()
      .setHorizontalAlignment("center")
  }

  /**
   * createCellsForTotalPaymentPerItem()
   * @param members 
   */
  private createCellsForTotalPaymentPerItem(members: Members) {
    const numMembers = members.getNumMembers();
    const columnIndexRatio = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2) + 1;
    const columnIndexPayment = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2) + 2;

    // Prepare values
    // 0th columns are for ratio, 1st columns are for payment
    const sumTable: string[][] = Util.createTable(DetailsSheet.NUM_INPUT_ROWS, 2);

    // Create cells for sum of each row
    for (let i = 0; i < DetailsSheet.NUM_INPUT_ROWS; i++) {
      const rowIndex = DetailsSheet.NUM_HEADER_ROWS + i + 1;
      const headStartCellStr = `R${DetailsSheet.ROW_INDEX_HEAD}C${DetailsSheet.NUM_LEFT_FIXED_COLUMNS + 1}`;
      const headEndCellStr = `R${DetailsSheet.ROW_INDEX_HEAD}C${DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2)}`;
      const sumStartCellStr = `R${rowIndex}C${DetailsSheet.NUM_LEFT_FIXED_COLUMNS + 1}`;
      const sumEndCellStr = `R${rowIndex}C${DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (numMembers * 2)}`;

      const ratioFormula =
        `=SUMIF(${headStartCellStr}:${headEndCellStr},"${DetailsSheet.TITLE_RATIO}",${sumStartCellStr}:${sumEndCellStr})`;
      const paymentFormula =
        `=SUMIF(${headStartCellStr}:${headEndCellStr},"${DetailsSheet.TITLE_PAYMENT}",${sumStartCellStr}:${sumEndCellStr})`;

      sumTable[i][0] = ratioFormula;
      sumTable[i][1] = paymentFormula;
    }

    // Write values to spreadsheet
    this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + 1, columnIndexRatio, DetailsSheet.NUM_INPUT_ROWS, 2)
      .setValues(sumTable)
      .setBackground(Util.BGCOLOR_SUM);

    this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + 1, columnIndexPayment, DetailsSheet.NUM_INPUT_ROWS, 1)
      .setNumberFormat(Util.FORMAT_CURRENCY)

    // Create cell for total sum
    const sumStartCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + 1}C${columnIndexPayment}`;
    const sumEndCellStr = `R${DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS}C${columnIndexPayment}`;
    const sumFormula = `SUM(${sumStartCellStr}:${sumEndCellStr})`;
    this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, columnIndexPayment)
      .setFormulaR1C1(sumFormula)
      .setBackground(Util.BGCOLOR_SUM);

    // Empty cell, but just set color
    this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + DetailsSheet.NUM_INPUT_ROWS + 1, columnIndexRatio)
      .setBackground(Util.BGCOLOR_SUM);

    // Set conditional format
    this.setTotalRatioCondition(columnIndexRatio);
    this.setTotalPaymentCondition(columnIndexPayment);
  }

  /**
   * setTotalRatioCondition()
   * @param columnIndex 
   */
  private setTotalRatioCondition(columnIndex: number) {
    const range = this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + 1, columnIndex, DetailsSheet.NUM_INPUT_ROWS, 1)
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberNotEqualTo(1)
      .setFontColor(Util.FONTCOLOR_ERROR)
      .setRanges([range])
      .build();
    
    const rules = this.sheet.getConditionalFormatRules()
    rules.push(rule);
    this.sheet.setConditionalFormatRules(rules);
  }

  /**
   * setTotalPaymentCondition()
   * @param columnIndex 
   */
  private setTotalPaymentCondition(columnIndex: number) {
    const rules = this.sheet.getConditionalFormatRules();

    // The condition also set to overall total payment, so the range is until NUM_INPUT_ROWS + 1
    for (let i = 0; i < DetailsSheet.NUM_INPUT_ROWS + 1; i++) {
      const rowIndex = DetailsSheet.NUM_HEADER_ROWS + i + 1;
      const refPrice = this.sheet.getRange(rowIndex, DetailsSheet.COLUMN_INDEX_PRICE).getA1Notation();
      const refSum = this.sheet.getRange(rowIndex, columnIndex).getA1Notation();
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=NE(${refSum},${refPrice})`)
        .setFontColor(Util.FONTCOLOR_ERROR)
        .setRanges([this.sheet.getRange(rowIndex, columnIndex)])
        .build();
      rules.push(rule);
    }

    this.sheet.setConditionalFormatRules(rules);
  }

  /**
   * getBuyerRange()
   * @returns 
   */
  getBuyerRange(): GoogleAppsScript.Spreadsheet.Range {
    return this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + 1, DetailsSheet.COLUMN_INDEX_BUYER, DetailsSheet.NUM_INPUT_ROWS, 1);
  }

  /**
   * getPaymentRange()
   * @param memberIndex 
   * @returns 
   */
  getPaymentRange(memberIndex: number): GoogleAppsScript.Spreadsheet.Range {
    const columnIndex = DetailsSheet.NUM_LEFT_FIXED_COLUMNS + (memberIndex * 2)
    return this.sheet.getRange(DetailsSheet.NUM_HEADER_ROWS + 1, columnIndex, DetailsSheet.NUM_INPUT_ROWS, 1);
  }

  /**
   * initialize()
   */
  initialize() {
    const membersSheet = new MembersSheet();
    const members = new Members(membersSheet);
    if (members.getNumMembers() < 2) {
      throw new Error("2名以上の参加者を入力してください");
    }
    this.sheet.clear();
    this.createLeftHeader();
    this.createDropDown(members);
    this.createHeader(members);
    this.createRatioInputForm(members);
    this.createCellsForTotalPrice();
    this.createCellsForTotalPaymentPerMember(members);
    this.createHeaderForTotalPaymentPerItem(members);
    this.createCellsForTotalPaymentPerItem(members);
  }

  /**
   * clear()
   */
  clear() {
    this.sheet.clear();
  }
}