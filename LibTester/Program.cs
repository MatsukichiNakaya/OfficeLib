//#define OUTLOOK
#define EXCEL

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
#if EXCEL
            using (var workBook = WorkBook.GetInstance(@".\WorkBook.xlsx"))
            {
                if(workBook == null) { return; }

                workBook.SelectSheet("sheet1");

                //// セル一つだけ
                //workBook.SetBackgroundColor(target: new Address("B4"),
                //                            color:  new Color(0, 0, 255));
                //// 複数セル一括
                //workBook.SetBackgroundColor(start: new Address("D3"),
                //                            end:   new Address("E5"),
                //                            color: new Color(0, 255, 0));

                //workBook.SetBorder(new Range("D3:E5"), new Thickness(new Border()));

                workBook.CopySheet("sheet1", "sheet4");
                workBook.Save();
            }
            Console.ReadLine();
#endif

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
        }
    }
}
