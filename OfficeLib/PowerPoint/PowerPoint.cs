//
// Work in progress
// 
using System;

namespace OfficeLib.PPT
{
    /// <summary>PowerPoint class</summary>
    public class PowerPoint : OfficeCore
    {
        /// <summary>Application object ID</summary>
        protected static readonly String PROG_ID = "PowerPoint.Application";
        /// <summary>Presentations object ID</summary>
        protected static readonly String OBJECT_PRESENTATIONS = "Presentations";


        #region --- Properties ---
        /// <summary>Presentation object</summary>
        public Object Presentation { get; private set; }



        #endregion

        #region --- Constructor ---
        /// <summary>Constructor</summary>
        public PowerPoint() : base(PROG_ID) { }
        #endregion

        #region --- Open ---
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public override Boolean Open(String filePath)
        {
            try
            {
                this.Path = System.IO.Path.GetFullPath(filePath);   // Must full path

                // Create instance of Application
                CreateApplication();
                // Disable warning indication. Prevent program from stopping with warning
                this.Application.SetProperty(PROP_DISP_ALERT,
                                             new Object[] { MsoTriState.msoFalse });

                // Get Presentations object from Application
                this.WorkArea = this.Application.GetProperty(OBJECT_PRESENTATIONS, null);
                
                // Get Presentation object from Presentations 
                this.Presentation = this.WorkArea.Method(METHOD_OPEN,
                                        new Object[] { this.Path,               // File path
                                                       MsoTriState.msoFalse,    // Readable and Writable
                                                       MsoTriState.msoFalse,    // Title is filename
                                                       MsoTriState.msoFalse }); // Hidden
            }
            catch (Exception ex)
            {
                //this.LastErrorLog = String.Format("Open:{0}\r\n{1}",
                //                                    ex.Message, ex.StackTrace);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public Boolean New(String filePath)
        {
            try
            {

            }
            catch (Exception ex)
            {
                //this.LastErrorLog = String.Format("Open:{0}\r\n{1}",
                //                                    ex.Message, ex.StackTrace);
                return false;
            }
            return true;
        }
        #endregion

        #region --- Save ---
        /// <summary>Save</summary>
        public void Save()
        {
            this.WorkArea.Method(METHOD_SAVE, null);
        }

        /// <summary>Save As</summary>
        /// <param name="savePath">Seve path</param>
        public void SaveAs(String savePath)
        {
            this.WorkArea.Method(METHOD_SAVE_AS, new Object[] { System.IO.Path.GetFullPath(savePath) });
        }
        #endregion

        #region --- Close ---
        /// <summary>
        /// Close
        /// </summary>
        public override void Close()
        {
            try
            {
                ClosePresentation();
                QuitAplication();
            }
            catch (Exception ex)
            {
                //this.LastErrorLog = String.Format("Close:{0}\r\n{1}",
                //                                     ex.Message, ex.StackTrace);
            }
        }

        /// <summary>
        /// Close Presentation
        /// </summary>
        private void ClosePresentation()
        {
            if (this.Presentation != null)
            {
                this.Presentation.Method(METHOD_CLOSE, null);
                ReleaseObject(this.Presentation);
            }
            ReleaseObject(this.WorkArea);
        }
        #endregion
    }
}
