﻿using System;

namespace OfficeLib
{
    /// <summary>
    /// Specifies whether to search the current directory,
    /// or the current directory and all subdirectories.
    /// </summary>
    public enum FolderSearchOption : Int32
    {
        /// <summary>
        /// Includes the current directory and all its subdirectories
        /// in a search operation.
        /// This option includes reparse points such as mounted drives
        /// and symbolic links in the search.
        /// </summary>
        AllFolders,
        /// <summary>
        /// Includes only the current directory in a search operation.
        /// </summary>
        TopFolderOnly,
    }

    /// <summary>Office Boolean</summary>
    public enum MsoTriState : Int32
    {
        /// <summary>Not supported</summary>
        msoTriStateToggle = -3,
        /// <summary>Not supported</summary>
        msoTriStateMixed = -2,
        /// <summary>True</summary>
        msoTrue = -1,
        /// <summary>False</summary>
        msoFalse = 0,
        /// <summary>Not supported</summary>
        msoCTrue = 1,
    }
}
