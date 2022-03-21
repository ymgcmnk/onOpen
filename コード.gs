/**
 * 更新日　2021/12/23
 * 更新者　XXXXXX
 * 
 * シンプルトリガー　onOpen()
 * 
 * ＜プロジェクト概要＞
 * スプレッドシートを開いたときに自動的に以下を実行する
 * ・独自メニューをシートに追加する
 * ・スプレッドシートの最終行にセルを移動
 * 
 * 関連note
 * https://note.com/0375/n/nff37c72fcc24
 * 
 */
function onOpen() {
  // const ui = SpreadsheetApp.getUi();
  // ui.alert('★注意★\n\n現在、スクリプトを改修中です。改修終了まで、このシートの使用は控えてください。\nご協力お願いいたします。', ui.ButtonSet.OK);

  addMenu();
  moveToLastRow();
}

/**
 * ＜function概要＞
 * 独自メニューをシートに表示する　汎用性、使いまわしの観点から、base　class　が良いか。
 * https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet?hl=en#addmenuname,-submenus
 * https://developers.google.com/apps-script/reference/base/menu?hl=en
*/
function addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('My Menu')
    .addItem('01_My Menu Item', 'myFunction01')
    .addItem('02_My Menu Item', 'myFunction02')
    .addItem('03_My Menu Item', 'myFunction03')
    .addItem('04_My Menu Item', 'myFunction04')
    .addSeparator()
    .addItem('05_My Menu Item', 'myFunction05"')
    .addItem('06_My Menu Item', 'myFunction06')
    .addItem('07_My Menu Item', 'myFunction07')
//       .addSubMenu(SpreadsheetApp.getUi().createMenu('My Submenu')
//           .addItem('One Submenu Item', 'mySecondFunction')
//           .addItem('Another Submenu Item', 'myThirdFunction'))
    .addToUi();
}




/**
 * ＜function概要＞
 * スプレッドシートの最終行にセルを移動
 * 
 * 参考URL
 * https://note.com/kawamura_/n/nb1865dfb3c77
 * 
 * セルが空白でもとにかく最終行に移動するときは　getMaxRows()　を使用
*/
function moveToLastRow() {
  //  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name');
  console.log(sheet.getName());
  const lastRow = sheet.getRange(1, 7).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  console.log(lastRow);
  sheet.getRange(lastRow + 1, 1).activate();
}

//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getActiveSheet();
//   const lastRow = sheet.getMaxRows();
//   sheet.getRange(lastRow, 1).activate();
