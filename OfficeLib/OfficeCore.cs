using System;
using System.Runtime.InteropServices;

namespace OfficeLib
{
    /// <summary>Microsoft Office application class</summary>
    /// <remarks>Excel, Word, PowerPoint</remarks>
    public abstract class OfficeCore : IDisposable
    {
        /// <summary>Command of Shapes</summary>
        protected const String OBJECT_SHAPES = "Shapes";

        /// <summary>Command of Open</summary>
        protected const String METHOD_OPEN = "Open";
        /// <summary>Command of Save</summary>
        protected const String METHOD_SAVE = "Save";
        /// <summary>Command of Save as</summary>
        protected const String METHOD_SAVE_AS = "SaveAs";
        /// <summary>Command of Close</summary>
        protected const String METHOD_CLOSE = "Close";
        /// <summary>Command of Quit</summary>
        protected const String METHOD_QUIT = "Quit";
        /// <summary>Command of Add</summary>
        protected const String METHOD_ADD = "Add";

        /// <summary>Command of Copy</summary>
        protected const String COMMAND_COPY = "Duplicate";
        /// <summary>Command of Cut</summary>
        protected const String COMMAND_CUT = "Cut";
        /// <summary>Command of Paste</summary>
        protected const String COMMAND_PASTE = "Paste";

        /// <summary>Version</summary>
        protected const String PROP_VERSION = "Version";
        /// <summary>count</summary>
        protected const String PROP_COUNT = "Count";
        /// <summary>Item</summary>
        protected const String PROP_ITEM = "Item";
        /// <summary>Items</summary>
        protected const String PROP_ITEMS = "Items";
        /// <summary>Name</summary>
        protected const String PROP_NAME = "Name";
        /// <summary>Left</summary>
        protected const String PROP_LEFT = "Left";
        /// <summary>Top</summary>
        protected const String PROP_TOP = "Top";
        /// <summary>Path</summary>
        protected const String PROP_PATH = "Path";
        /// <summary>Saved</summary>
        protected const String PROP_SAVED = "Saved";
        /// <summary>Display alert</summary>
        protected const String PROP_DISP_ALERT = "DisplayAlerts";

        /// <summary>Application ID</summary>
        public String ApplicationID { get; }

        /// <summary>File path of this application</summary>
        /// <remarks>
        /// Must set full path
        /// </remarks>
        public String Path { get; protected set; }

        /// <summary>Application resource</summary>
        protected Object Application { get; set; }

        /// <summary>Work area</summary>
        protected Object WorkArea { get; set; }

        /// <summary>Version of Application</summary>
        public String Version
        {
            get { return this.Application.GetProperty(PROP_VERSION, null) as String; }
        }

        /// <summary>Constructor</summary>
        /// <param name="appID">Application ID</param>
        public OfficeCore(String appID)
        {
            this.ApplicationID = appID;
        }

        #region --- Open ---
        /// <summary>
        /// Create Application
        /// </summary>
        /// <returns></returns>
        protected virtual Boolean CreateApplication()
        {
            if (this.Application != null) { Close(); }

            var appType = Type.GetTypeFromProgID(this.ApplicationID);
            this.Application = Activator.CreateInstance(appType);
            // Null failed to create
            return this.Application != null;
        }

        /// <summary>Open Application</summary>
        /// <param name="filePath">Application path</param>
        /// <returns>Success(true) or failure(false) of execution</returns>
        public abstract Boolean Open(String filePath);
        #endregion

        #region --- Close ---
        /// <summary> Dispose </summary>
        /// <remarks>Own 'Close' method calling.</remarks>
        public void Dispose() => Close();

        /// <summary>Close aplication</summary>
        /// <remarks>
        /// It must be done at the end.
        /// If not called, the process will keep capturing the file.
        /// (Especially Excel!)
        /// </remarks>
        public abstract void Close();

        /// <summary>
        /// Resource release
        /// </summary>
        /// <param name="target">Target object</param>
        protected void ReleaseObject(Object target)
        {
            try
            {
                if (target == null) { return; }
                // Free the object's resources.
                do { } while (Marshal.ReleaseComObject(target) > 0);
            }
            finally { target = null; }
        }

        /// <summary>
        /// Resources release
        /// </summary>
        /// <param name="targets">List of target object</param>
        protected void ReleaseObjects(params Object[] targets)
        {
            foreach (var target in targets)
            {
                ReleaseObject(target);
            }
        }

        /// <summary>Quit Application</summary>
        protected void QuitAplication()
        {
            this.Application?.Method(METHOD_QUIT, null);
            ReleaseObject(this.Application);
        }
        #endregion
    }
}