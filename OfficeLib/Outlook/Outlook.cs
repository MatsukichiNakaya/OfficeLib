using System;
using static OfficeLib.Commands;

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

        /// <summary>Indicates the recipient type for the Item.</summary>
        public enum OlMailRecipientType
        {
            /// <summary>Originator (sender) of the Item.</summary>
            olOriginator = 0,
            /// <summary>The recipient is specified in the To property of the Item.</summary>
            olTo,
            /// <summary>The recipient is specified in the CC property of the Item.</summary>
            olCC,
            /// <summary>The recipient is specified in the BCC property of the Item.</summary>
            olBCC,
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
        /// Get Folder object
        /// </summary>
        /// <param name="folderType"></param>
        /// <returns></returns>
        public Object GetFolder(OlDefaultFolders folderType)
        {
            Object ns = null;
            Object folder = null;

            try
            {
                ns = this.Application.Method(METHOD_GET_NAMESPACE, new Object[] { ARG_MAPI });
                folder = ns.Method(METHOD_GET_DEFAULT_FOLDER, new Object[] { folderType });
            }
            catch { ReleaseObject(folder); }
            finally { ReleaseObject(ns); }

            return folder;
        }

        /// <summary>
        /// Get child folder object
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="childFolderName"></param>
        /// <returns></returns>
        public Object GetChildFolder(Object folder, String childFolderName)
        {
            Object specFolder = null;
            try
            {
                specFolder = folder.GetProperty(PROP_FOLDERS, new Object[] { childFolderName });
            }
            catch { ReleaseObject(specFolder); }

            return specFolder;
        }

        /// <summary>
        /// Retrieve mail from specified folder.
        /// </summary>
        /// <param name="folder">Folder object</param>
        /// <returns></returns>
        public EMail[] GetMails(Object folder)
        {
            EMail[] result = null;
            Object items = null;
            try
            {
                items = folder.GetProperty(PROP_ITEMS);
                // Get number of items and initialize array.
                result = new EMail[items.GetProperty(PROP_COUNT).To<Int32>()];

                System.Threading.Tasks.Parallel.For(0, result.Length, (no) =>
                {
                    Object item = items.GetProperty(PROP_ITEM, new Object[] { no + 1 });
                    result[no] = new EMail(item);
                    ReleaseObject(item);
                });
            }
            finally
            {   // Todo : Is releasable [folder]?
                ReleaseObjects(folder, items);
            }
            return result;
        }

        #region Unread Count Method
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
        /// Get the number of unread counts in specified folder
        /// </summary>
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
                    // vba is 1 origin
                    result += innerFolder.GetProperty(PROP_UNREAD_ITEM_COUNT).To<Int32>();

                    // Is there a lower hierarchy
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
        #endregion

        /// <summary>
        /// Add folder from specified folder.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="folderName"></param>
        /// <param name="folderType"></param>
        public void AddFolder(Object folder, String folderName, OlDefaultFolders folderType)
        {
            folder.GetProperty(PROP_FOLDERS).Method(METHOD_ADD, new Object[] { folderName, folderType });
        }

        /// <summary>
        /// Auto Filtering
        /// </summary>
        /// <param name="srcFolder">Src folder object</param>
        /// <param name="destFolder">Dest folder object</param>
        /// <param name="filter">Filter script</param>
        public void AutoFiltering(Object srcFolder, Object destFolder,
                                  Func<Object, Boolean> filter)
        {
            Object items = null;
            try
            {
                items = srcFolder.GetProperty(PROP_ITEMS);
                Int32 mailCount = items.GetProperty(PROP_COUNT).To<Int32>();

                for (var no = 0; no < mailCount; no++)
                {
                    Object item = items.GetProperty(PROP_ITEM, new Object[] { no + 1 });
                    if (filter(item))
                    {
                        item.Method(METHOD_MOVE, new Object[] { destFolder });
                        // Return step.
                        no -= 1;
                        // param reset and restart.
                        ReleaseObject(items);
                        items = srcFolder.GetProperty(PROP_ITEMS);
                        mailCount = items.GetProperty(PROP_COUNT).To<Int32>();
                    }
                    ReleaseObject(item);
                }
            }
            finally { ReleaseObject(items); }
        }
    }
}
