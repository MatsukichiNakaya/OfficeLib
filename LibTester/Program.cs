using OfficeLib;        // Dll ベース部分
using OfficeLib.EML;    // Outlook
using OfficeLib.PPT;    // PowerPoint
using OfficeLib.XLS;    // Excel

using System;
using System.IO;
using System.Linq;
using System.Text;

using static OfficeLib.EnumSheetPermission;

namespace LibTester
{
    /// <summary>
    /// テストプログラム
    /// </summary>
    class Program
    {
        static void Main(String[] args)
        {
            var wb = new WorkBook("WorkBook.xlsx");
            wb.AddSheet(new Sheet1());

            wb.ReadPreset();

            var sh = wb[Sheet1.SHEET_NAME];

            var table = new String[4][];
            for (var r = 0; r < table.Length; r++)
            {
                table[r] = new String[4];
                for (var c = 0; c < table[r].Length; c++)
                {
                    table[r][c] = (r + c).ToString();
                }
            }

            sh[new Address("B2"), new Address("D4")] = table;


            Console.ReadLine();



#if OUTLOOK
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
#endif
        }

        [ExcelSheet(ReadWrite)]
        class Sheet1 : WorkSheet
        {
            public const String SHEET_NAME = "Sheet1";

            public Sheet1() : base(SHEET_NAME, "E5") { }

            public override void Read(Excel excel)
            {
                base.Read(excel);
            }

            public override void Write(Excel excel)
            {
                base.Write(excel);
            }
        }
    }
}
