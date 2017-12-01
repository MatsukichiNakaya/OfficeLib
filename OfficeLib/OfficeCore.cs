using System;
using System.Runtime.InteropServices;
using static OfficeLib.Commands;

namespace OfficeLib
{
    /// <summary>Microsoft Office application class</summary>
    /// <remarks>Excel, Word, PowerPoint</remarks>
    public abstract class OfficeCore : IDisposable
    {
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
        public static void ReleaseObject(Object target)
        {
            try
            {
                if (target == null) { return; }
                // Free the object's resources.
                do { } while (0 < Marshal.ReleaseComObject(target));
            }
            finally { target = null; }
        }

        /// <summary>
        /// Resources release
        /// </summary>
        /// <param name="targets">List of target object</param>
        public static void ReleaseObjects(params Object[] targets)
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