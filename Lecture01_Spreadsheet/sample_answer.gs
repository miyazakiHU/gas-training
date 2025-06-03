function processSpreadsheet() {
  const SPREADSHEET_ID = '<<ここに自分のコピーしたスプレッドシートIDを貼り付け>>';
  const SHEET_NAME = 'シート1'; // コピー元の名前そのままの場合

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  const header = values[0];
  const rows = values.slice(1);

  // 空行を削除（B列: 氏名が空白の行）
  for (let i = rows.length - 1; i >= 0; i--) {
    if (!rows[i][1] || rows[i][1].toString().trim() === '') {
      sheet.deleteRow(i + 2); // +2: ヘッダー分 + インデックス調整
    }
  }

  // 再取得（削除後のシート）
  const rangeAfterDelete = sheet.getDataRange();
  const allValues = rangeAfterDelete.getValues();
  const body = allValues.slice(1);

  // C列（誕生日）昇順ソート（列番号は0始まり→2がC列）
  sheet.getRange(2, 1, body.length, header.length)
    .sort({ column: 3, ascending: true });

  // 条件付き書式の設定（D列: 好感度スコア）
  const scoreColumn = 4;
  const numRows = sheet.getLastRow() - 1;

  const range = sheet.getRange(2, scoreColumn, numRows);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground('#c6efce') // 緑系
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(50)
      .setBackground('#ffc7ce') // 赤系
      .setRanges([range])
      .build(),
  ];

  const existingRules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules([...existingRules, ...rules]);

  Logger.log('スプレッドシートの整形が完了しました。');
}
