using OfficeLib;        // Dll ベース部分
using OfficeLib.EML;    // Outlook
using OfficeLib.PPT;    // PowerPoint
using OfficeLib.XLS;    // Excel

using System;
using System.IO;
using System.Linq;
using System.Text;

namespace LibTester
{
    /// <summary>
    /// テストプログラム
    /// </summary>
    class Program
    {
        private const String OUTPUT_SHEET = "Output";

        static void Main(String[] args)
        {

            //Field<String> GeneTable;

            //{   // Excel Bookの読込みと値の取得
            //    var workBook = new WorkBook(@"D:\Workspase\20170407PSO\PSO\Dat\"
            //                                + "AEMSシミュレータ入力シート.xlsx");

            //    //workBook.AddSheet(new WorkSheet(name: OUTPUT_SHEET, endAddress: "BL309"));
            //    workBook.AddSheet(new SheetCondition());

            //    // 定義されているシートを一括で読み込む
            //    workBook.ReadPreset();
            //    // すべて読み込む場合はこちら
            //    // ※シートを定義しなくてもファイル中のすべてのシートを読む
            //    // wb.Read();
            //    // シート一枚指定はこちら
            //    // wb.Read(OutputSheet);

            //    // 参照を間単にする為にテーブルを定義 ↓発電機 5 まで, 48時間分の範囲
            //    //GeneTable = workBook[OUTPUT_SHEET].GetTable(startAddrStr: "D213",
            //    //                                            endAddrStr: "AG308").Convert<String>();
            //}

            //// --- 値の処理 ---
            //// 時刻単位で値を取得
            //var readingField = new StringBuilder(2048);
            //for (var row = 0; row < GeneTable.Row; row++)
            //{   // (取得した項目を","で結合する ( テーブルの[i]行目.列インデックスを6の剰余が指定の数である項目を取得 )).末尾に改行
            //    readingField.Append(String.Join(",", GeneTable[row].Where((val, col) => col % 6 == OUTPUT))).Append("\r\n");
            //}

            //// 書き出し
            //Write(@".\reated.txt", readingField.ToString(), false, Encoding.UTF8);

            // Outlookテストコード
            {   // Mail
                // Outlookへの接続を行います
                var ol = new Outlook();
                ol.Connect();

                Int32 unreads = 0;// ol.GetUnreadCount("システム開発");
                unreads = ol.GetUnreadCount(FolderSearchOption.AllFolders);
                Console.WriteLine(unreads);

                // ディレクトリの一覧取得
                //var folders = ol.GetInboxFolders();
                //foreach (var f in folders)
                //{
                //    Console.WriteLine(f);
                //}
            }
            Console.ReadLine();
        }

        /// <summary>
        /// テキストの出力
        /// </summary>
        /// <param name="path">ファイルパス</param>
        /// <param name="text">テキストの内容</param>
        /// <param name="append">ファイルへの追加(true)または上書き(false)</param>
        /// <param name="enc">エンコード情報</param>
        private static void Write(String path, String text, Boolean append, Encoding enc)
        {
            using (var writer = new StreamWriter(path, append, enc))
            {
                writer.Write(text);
                writer.Flush();
            }
        }
    }

    class SheetCondition : WorkSheet
    {
        public SheetCondition() : base("Condition", "F21") { }

        public override void Read(Excel excel)
        {
            base.Read(excel);
            // 計算条件の読込み
            this.AddTable("Tbl_Condition", GetTable("F4", "F19"));
        }
    }

}
