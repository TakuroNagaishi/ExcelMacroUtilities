# ExcelMacroUtilities

Excel便利マクロ保管用

## Modules

### [FormatDocument](src/FormatBook.bas)

Bookの体裁を整える

Book内の全Sheetに対し以下操作を行う
- 選択セルをA1に変更
- 拡大率を100％に変更

その後先頭のSheetをアクティブにして終了

### [ListSheetNamesToSheet](src/ListSheetNamesToSheet.bas)

Book内の全Sheetの名前を一覧化する

新規Sheetを作成し、そこにBook内全Sheetの名前を出力する
