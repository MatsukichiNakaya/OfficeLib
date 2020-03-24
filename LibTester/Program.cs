//#define OUTLOOK
#define EXCEL
//#define POWERPOINT

using OfficeLib;        // Dll ベース部分
using OfficeLib.EML;    // Outlook
using OfficeLib.PPT;    // PowerPoint
using OfficeLib.XLS;    // Excel

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
            #region Excel
#if EXCEL
            /*
            //using (var workBook = WorkBook.GetInstance(@".\WorkBook.xlsx"))
            //{
            //    if (workBook == null) { return; }

            //    workBook.SelectSheet("sheet1");

            //    //// セル一つだけ
            //    //workBook.SetBackgroundColor(target: new Address("B4"),
            //    //                            color:  new Color(0, 0, 255));
            //    //// 複数セル一括
            //    //workBook.SetBackgroundColor(start: new Address("D3"),
            //    //                            end:   new Address("E5"),
            //    //                            color: new Color(0, 255, 0));

            //    //workBook.SetBorder(new Range("D3:E5"), new Thickness(new Border()));

            //    //workBook.CopySheet("sheet1", "sheet4");
            //    //workBook.SetSheetProperty("sheet3", "Visible", new Object[] { OfficeLib.MsoTriState.msoTrue });
            //    //workBook.MoveSheet("Sheet1", beforeSheetName: "Sheet3");
            //    //workBook.Save();
            //}
            */
            /*
            var workBook = new WorkBook(@".\WorkBook.xlsx");

            //workBook.AddSheet(new Sheet1());

            //workBook.WriteSheet("Sheet1");

            workBook.Read("Sheet2", XlGetValueFormat.xlFormula);

            Object val = workBook["Sheet2"][new Address("B3")];
            */
            //*
            using (var book = WorkBook.GetInstance(@".\WorkBook.xlsx", isAutoSave:true))
            {

                //book.RemoveSheet("Sheet3");

                //var val = book.GetCellValue("A4", "A4", XlGetValueFormat.xlValue);
                //Console.WriteLine(val);
                //var set = new String[1, 1];
                //set[0, 0] = "3";
                //book.SetCellValue(set, "A1", XlGetValueFormat.xlValue);
                //var val2 = book.GetCellValue("A4", "A4", XlGetValueFormat.xlFormula);
                //Console.WriteLine(val2);
                //var ret = book.GetLastCell();

                //using (var sheet = book.GetSheet("Sheet2"))
                //{
                //    Excel.Macro(sheet.ComObject, "Move", null);
                //}

                //book.SetCellValue(null, "C5", XlGetValueFormat.xlFormula);


                //book.CellCopy("Sheet2", "A6", "A6");

                //using (var book2 = WorkBook.GetInstance(@".\WorkBook2.xlsx")) {
                //    book2.CellPaste("Sheet2", "B6", "B6");
                //    book2.Save();
                //}
                //book.AtherBookCellPaste(@".\WorkBook2.xlsx", "Sheet2", "B6", "B6");

                book.SelectSheet("Sheet2");
                // book.AddChart();

                //using (var sheet = book.GetSheet("Sheet2")) {

                //    var count = sheet.Method("ChartObjects").GetProperty("Count").ToInt();

                //    Console.WriteLine(count);
                //}
                //

                //book.SetChartType(1, XlChartType.xlLine);

                book.SetChartTypeSeries(1, new XlChartType[] { XlChartType.xlLine, XlChartType.xlColumnClustered });
            }//*/

            Console.WriteLine("Complete");
            Console.ReadLine();
#endif
#endregion


#if OUTLOOK
            // Outlookテストコード
            {   // Mail
                // Outlookへの接続を行います
                var ol = new Outlook();
                ol.Connect();

                // 送信済みメール一覧
                //EMail[] mails = ol.GetMails(ol.GetFolder(Outlook.OlDefaultFolders.olFolderSentMail));

                // 未読メール件数（全件取得）
                //Int32 count = ol.GetUnreadCount(FolderSearchOption.AllFolders);
                //Console.WriteLine(count);
                //Console.ReadLine();

                //// フォルダの追加
                //Object folder = ol.GetFolder(Outlook.OlDefaultFolders.olFolderInbox);
                //ol.AddFolder(folder, "重要", Outlook.OlDefaultFolders.olFolderInbox);
                //OfficeCore.ReleaseObject(folder);

                // 振り分け処理マクロ作成
                Object srcfolder = ol.GetFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Object destfolder = ol.GetChildFolder(srcfolder, "三洋");

                ol.AutoFiltering(srcfolder, destfolder, (mail) =>
                {   // 送信元の情報が指定のパターンの場合にマッチしているか？
                    return Regex.IsMatch(new EMail(mail).From.Address, @"@jp\.panasonic\.com");
                });
            }
#endif

#if POWERPOINT
            using (var pp = new PowerPoint())
            {
                pp.Open(@"presentation.pptx");

                pp.SelectSlide(2);
            }

            Console.WriteLine();
#endif
        }
    }

    [ExcelSheet(EnumSheetPermission.ReadWrite)]
    class Sheet1 : WorkSheet
    {
        public Sheet1() : base("Sheet1", "F5") { }

        public override void Write(Excel excel)
        {
            excel.SelectSheet(this.Name);

            SetValue(excel,
                     "=(TIMEVALUE(\"20:00\")-TIMEVALUE(\"18:30\"))*24",
                     new Address("A1"),
                     XlGetValueFormat.xlFormula);
        }
    }
}
