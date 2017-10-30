            
一番簡単な構成
	    const String SHEET_SOFT_LIST = "SoftwareList";
            
            /*** 既存のブック、既存のシートを読む場合 ***/
            var book = new WorkBook(@".\SoftwareManagement.xlsx");
            book.AddSheet(new WorkSheet(SHEET_SOFT_LIST));

            // シート読込み
            book.ReadBook(SHEET_SOFT_LIST);

            // データがあるとする
            var list = new String[] { "notepad", "paint", "calculator" };

            // 参照セル作成
            var currentCell = new Range("B2");

            foreach (var item in list)  // B2 ～ B4に値を書き込む
            {
                book[SHEET_SOFT_LIST][currentCell] = item;
                currentCell.Shift(0, 1);   // 一つ下へ移動
            }

            // 変更を保存
            book.WriteBook();


内部のメモリ割当量を減らす工夫
        static void Main(string[] args)
        {
            /*** 既存のブック、既存のシートを読む場合 ***/
            var book = new WorkBook(@".\SoftwareManagement.xlsx");
            book.AddSheet(new SoftListSheet());

            // シート読込み
            book.ReadBook(SheetName.SOFT_LIST);

            // データがあるとする
            var list = new String[] { "notepad", "paint", "calculator" };

            // 参照セル作成
            var currentCell = new Range("B2");

            foreach (var item in list)  // B2 ～ B4に値を書き込む
            {
                book[SheetName.SOFT_LIST][currentCell] = item;
                currentCell.Shift(0, 1);   // 一つ下へ移動
            }

            // 変更を保存
            book.WriteBook();
        }

        // 共通の参照
        class SheetName
        {
            public const String SOFT_LIST = "SoftwareList";
        }
    
        // 定義無の状態だと1000 x 1000の範囲を設定するのでこの定義でコストを節約する
        [ExcelSheet(EnumSheetPermission.ReadWrite, ColMax = 20, RowMax = 20)]
        class SoftListSheet : WorkSheet
        {
            public SoftListSheet() : base(SheetName.SOFT_LIST) { }
        }