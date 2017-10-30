using System;

namespace OfficeLib.EML
{
    /// <summary>
    /// 
    /// </summary>
    public class Outlook : OfficeCore
    {
        /// <summary>Application object ID</summary>
        protected const String PROG_ID = "Outlook.Application";

        /// <summary>Command of GetNamespace</summary>
        protected const String METHOD_GET_NAMESPACE = "GetNamespace";
        /// <summary>Command of GetNamespace</summary>
        protected const String METHOD_GET_DEFAULT_FOLDER = "GetDefaultFolder";
        /// <summary>Command of Restrict</summary>
        protected const String METHOD_RESTRICT = "Restrict";

        /// <summary>Unread item count</summary>
        protected const String PROP_UNREAD_ITEM_COUNT = "UnReadItemCount";
        /// <summary>Folders</summary>
        protected const String PROP_FOLDERS = "Folders";
        /// <summary></summary>
        protected const String PROP_SUBJECT = "Subject";

        /// <summary></summary>
        protected const String ARG_MAPI = "MAPI";

        /// <summary>Specifies the folder type for a specified folder</summary>
        public enum OlDefaultFolders
        {   /// <summary>The Calendar folder</summary>
            olFolderCalendar = 9,
            /// <summary>The Conflicts folder (subfolder of the Sync Issues folder)  Only available for an Exchange account</summary>
            olFolderConflicts = 19,
            /// <summary>The Contacts folder</summary>
            olFolderContacts = 10,
            /// <summary>The Deleted Items folder</summary>
            olFolderDeletedItems = 3,
            /// <summary>The Drafts folder</summary>
            olFolderDrafts = 16,
            /// <summary>The Inbox folder</summary>
            olFolderInbox = 6,
            /// <summary>The Journal folder</summary>
            olFolderJournal = 11,
            /// <summary>The Junk E-Mail folder</summary>
            olFolderJunk = 23,
            /// <summary>The Local Failures folder (subfolder of the Sync Issues folder)  Only available for an Exchange account</summary>
            olFolderLocalFailures = 21,
            /// <summary>The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account</summary>
            olFolderManagedEmail = 29,
            /// <summary>The Notes folder</summary>
            olFolderNotes = 12,
            /// <summary>The Outbox folder</summary>
            olFolderOutbox = 4,
            /// <summary>The Sent Mail folder</summary>
            olFolderSentMail = 5,
            /// <summary>The Server Failures folder (subfolder of the Sync Issues folder)  Only available for an Exchange account</summary>
            olFolderServerFailures = 22,
            /// <summary>The Suggested Contacts folder</summary>
            olFolderSuggestedContacts = 30,
            /// <summary>The Sync Issues folder. Only available for an Exchange account</summary>
            olFolderSyncIssues = 20,
            /// <summary>The Tasks folder</summary>
            olFolderTasks = 13,
            /// <summary>The To Do folder</summary>
            olFolderToDo = 28,
            /// <summary>The All Public Folders folder in the Exchange Public Folders store.Only available for an Exchange account</summary>
            olPublicFoldersAllPublicFolders = 18,
            /// <summary>The RSS Feeds folder</summary>
            olFolderRssFeeds = 25,
        }

        /// <summary>
        /// 
        /// </summary>
        public Outlook() : base(PROG_ID) { }

        /// <summary>
        /// Open E-Mail
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public override Boolean Open(String filePath)
        {
            return false;
        }

        /// <summary>
        /// Connect to Outlook
        /// </summary>
        /// <returns></returns>
        public Boolean Connect()
        {
            return base.CreateApplication();
        }

        /// <summary>
        /// Close Outlook
        /// </summary>
        public override void Close()
        {
            base.QuitAplication();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public Int32 GetItemCount()
        {
            return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public String[] GetInboxFolders()
        {
            String[] result = null;
            Object ns = null;
            Object inbox = null;
            Object folders = null;
            try
            {
                ns = this.Application.Method(METHOD_GET_NAMESPACE, new Object[] { ARG_MAPI });
                // Get Folders Property
                inbox = ns.Method(METHOD_GET_DEFAULT_FOLDER, new Object[] { OlDefaultFolders.olFolderInbox });
                folders = inbox.GetProperty(PROP_FOLDERS);
                Int32 folderCount = folders.GetProperty(PROP_COUNT).To<Int32>();

                result = new String[folderCount];
                for (Int32 i = 0; i < result.Length; i++)
                {   // Get Name property
                    result[i] = inbox.GetProperty(PROP_FOLDERS, new Object[] { i + 1 })
                                     .GetProperty(PROP_NAME).ToString();
                }
            }
            finally
            {
                // Free the Objects
                ReleaseObjects(ns, inbox, folders);
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public EMail[] GetInboxMails()
        {
            Object ns = null;
            Object inbox = null;

            try
            {   // Get Namespace Object
                ns = this.Application.Method(METHOD_GET_NAMESPACE, new Object[] { ARG_MAPI });
                // Get Folders Property
                inbox = ns.Method(METHOD_GET_DEFAULT_FOLDER, new Object[] { OlDefaultFolders.olFolderInbox });


                return GetMails(inbox);
            }
            finally
            {   // Free the Objects
                ReleaseObjects(ns, inbox);
            }
        }

        /// <summary>
        /// Todo : Mail 
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        private EMail[] GetMails(Object folder)
        {
            Object folders = null;
            Object innerFolder = null;
            Int32 innerCount = 0;
            var result = new System.Collections.Generic.List<EMail>();

            try
            {
                folders = folder.GetProperty(PROP_FOLDERS);
                Int32 folderCount = folders.GetProperty(PROP_COUNT).To<Int32>();
                Object items = null;
                Int32 itemCount = 0;
                
                for (var i = 0; i < folderCount; i++)
                {
                    innerFolder = folder.GetProperty(PROP_FOLDERS, new Object[] { i + 1 });
                    // vba は 配列が 1 オリジン
                    items = innerFolder.GetProperty(PROP_ITEMS);
                    itemCount = items.GetProperty(PROP_COUNT).To<Int32>();
                    var itemList = new Object[itemCount];
                    for (var no = 0; no < itemCount; no++)
                    {
                        itemList[no] = items.Method(PROP_ITEM, new Object[] { no + 1 });
                    }

                    // Todo : EMail クラスに変換する
                    // result.Add(); 

                    ReleaseObjects(itemList);

                    // 下の階層があるか
                    innerCount = innerFolder.GetProperty(PROP_FOLDERS)
                                            .GetProperty(PROP_COUNT).To<Int32>();
                    if (0 < innerCount)
                    {
                        result.AddRange(GetMails(innerFolder));
                    }
                    ReleaseObject(innerFolder);
                }
            }
            finally
            {
                ReleaseObjects(folders);
            }
            return result.ToArray();
        }

        /// <summary>
        /// Get the number of unread counts in inbox
        /// </summary>
        /// <returns>Unread mail count</returns>
        public Int32 GetUnreadCount(FolderSearchOption option)
        {
            Int32 result = 0;
            Object ns = null;
            Object inbox = null;
            try
            {
                // Get Namespace Object
                ns = this.Application.Method(METHOD_GET_NAMESPACE, new Object[] { ARG_MAPI });
                // Get Folders Property
                inbox = ns.Method(METHOD_GET_DEFAULT_FOLDER, new Object[] { OlDefaultFolders.olFolderInbox });
                // Get Item count. Convert to Int32
                result = inbox.GetProperty(PROP_UNREAD_ITEM_COUNT).To<Int32>();

                if (option == FolderSearchOption.AllFolders)
                {
                    result += UnreadCount(inbox);
                }
            }
            finally
            {   // Free the Objects
                ReleaseObjects(ns, inbox);
            }
            return result;
        }

        /// <summary>
        /// Get the number of unread counts in specified folder
        /// </summary>
        /// <param name="folderName">folder name</param>
        /// <returns>Unread mail count</returns>
        public Int32 GetUnreadCount(String folderName)
        {
            Int32 result = 0;
            Object ns = null;
            Object inbox = null;
            Object specFolder = null;
            try
            {
                // Get Namespace Object
                ns = this.Application.Method(METHOD_GET_NAMESPACE, new Object[] { ARG_MAPI });
                // Get Folder Object
                inbox = ns.Method(METHOD_GET_DEFAULT_FOLDER, new Object[] { OlDefaultFolders.olFolderInbox });
                // Get Folder Object
                specFolder = inbox.GetProperty(PROP_FOLDERS, new Object[] { folderName });
                // Get Item count. Convert to Int32
                result = inbox.GetProperty(PROP_UNREAD_ITEM_COUNT).To<Int32>();
            }
            finally
            {   // Free the Objects
                ReleaseObjects(ns, inbox, specFolder);
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private Int32 UnreadCount(Object folder)
        {
            Object folders = null;
            Object innerFolder = null;
            Int32 innerCount = 0;
            Int32 result = 0;

            try
            {
                folders = folder.GetProperty(PROP_FOLDERS);
                Int32 folderCount = folders.GetProperty(PROP_COUNT).To<Int32>();

                for (var i = 0; i < folderCount; i++)
                {
                    innerFolder = folder.GetProperty(PROP_FOLDERS, new Object[] { i + 1 });
                    // vba は 配列が 1 オリジン
                    result += innerFolder.GetProperty(PROP_UNREAD_ITEM_COUNT).To<Int32>();

                    // 下の階層があるか
                    innerCount = innerFolder.GetProperty(PROP_FOLDERS).GetProperty(PROP_COUNT).To<Int32>();
                    if (0 < innerCount)
                    {
                        result += UnreadCount(innerFolder);
                    }
                    ReleaseObject(innerFolder);
                }
            }
            finally
            {
                ReleaseObjects(folders);
            }
            return result;
        }
    }
}
