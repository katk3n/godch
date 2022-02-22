class Util {
  static readonly FONTCOLOR_INACTIVE = "gray";
  static readonly FONTCOLOR_ERROR = "red";
  static readonly BGCOLOR_INACTIVE = "#f5f5f5";
  static readonly BGCOLOR_MEMBER = "#e0ffff";
  static readonly BGCOLOR_SUM = "#fff2cc";
  static readonly FORMAT_CURRENCY = "[$¥-411]#,##0";

  static createTable(numRows: number, numColumns: number): string[][] {
    const table: string[][] = new Array(numRows);
    for (let i = 0; i < numRows; i++) {
      table[i] = new Array(numColumns).fill("");
    }

    return table;
  }
}