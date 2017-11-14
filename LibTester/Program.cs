#define OUTLOOK
//#define EXCEL

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
                Object destfolder = ol.GetChildFolder(srcfolder, "GMail");

                ol.AutoFiltering(srcfolder, destfolder, (mail) =>
                {   // 送信元の情報が指定のパターンの場合にマッチしているか？
                    return Regex.IsMatch(new EMail(mail).From.Address, @"@gmail\.com");
                });
            }
#endif
        }
    }
}
