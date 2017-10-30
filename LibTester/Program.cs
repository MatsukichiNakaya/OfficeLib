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
        static void Main(String[] args)
        {
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
    }
}
