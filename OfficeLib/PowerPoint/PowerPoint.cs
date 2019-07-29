//
// Work in progress
// 
using System;
using static OfficeLib.Commands;

namespace OfficeLib.PPT
{
    /// <summary>PowerPoint class</summary>
    public class PowerPoint : OfficeCore
    {
        /// <summary>Application object ID</summary>
        protected static readonly String PROG_ID = "PowerPoint.Application";
        /// <summary>Presentations object ID</summary>
        protected static readonly String OBJECT_PRESENTATIONS = "Presentations";

        /// <summary>Property Slides</summary>
        protected static readonly String PROP_SLIDES = "Slides";

        /// <summary>Property NotesPage</summary>
        protected static readonly String PROP_NOTESPAGE = "NotesPage";

        #region --- Properties ---
        /// <summary>Presentation object</summary>
        public Object Presentation { get; private set; }

        /// <summary>Slide object</summary>
        public Object Slide { get; private set; }

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
            catch (Exception) { throw; }
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
            catch (Exception) { throw; }
            return true;
        }
        #endregion

        #region --- Slides ---
        /// <summary>
        /// 
        /// </summary>
        /// <param name="slideNo"></param>
        public void SelectSlide(Int32 slideNo)
        {
            var slide = this.Presentation.GetProperty(PROP_SLIDES, new Object[] { slideNo });
            String name = slide.GetProperty(PROP_NOTESPAGE).ToString();


            ReleaseObject(slide);
            Console.WriteLine(name);
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
            catch (Exception) { throw; }
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
