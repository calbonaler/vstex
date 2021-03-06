/********************************************************************************************

Copyright (c) Microsoft Corporation 
All rights reserved. 

Microsoft Public License: 

This license governs use of the accompanying software. If you use the software, you 
accept this license. If you do not accept the license, do not use the software. 

1. Definitions 
The terms "reproduce," "reproduction," "derivative works," and "distribution" have the 
same meaning here as under U.S. copyright law. 
A "contribution" is the original software, or any additions or changes to the software. 
A "contributor" is any person that distributes its contribution under this license. 
"Licensed patents" are a contributor's patent claims that read directly on its contribution. 

2. Grant of Rights 
(A) Copyright Grant- Subject to the terms of this license, including the license conditions 
and limitations in section 3, each contributor grants you a non-exclusive, worldwide, 
royalty-free copyright license to reproduce its contribution, prepare derivative works of 
its contribution, and distribute its contribution or any derivative works that you create. 
(B) Patent Grant- Subject to the terms of this license, including the license conditions 
and limitations in section 3, each contributor grants you a non-exclusive, worldwide, 
royalty-free license under its licensed patents to make, have made, use, sell, offer for 
sale, import, and/or otherwise dispose of its contribution in the software or derivative 
works of the contribution in the software. 

3. Conditions and Limitations 
(A) No Trademark License- This license does not grant you rights to use any contributors' 
name, logo, or trademarks. 
(B) If you bring a patent claim against any contributor over patents that you claim are 
infringed by the software, your patent license from such contributor to the software ends 
automatically. 
(C) If you distribute any portion of the software, you must retain all copyright, patent, 
trademark, and attribution notices that are present in the software. 
(D) If you distribute any portion of the software in source code form, you may do so only 
under this license by including a complete copy of this license with your distribution. 
If you distribute any portion of the software in compiled or object code form, you may only 
do so under a license that complies with this license. 
(E) The software is licensed "as-is." You bear the risk of using it. The contributors give 
no express warranties, guarantees or conditions. You may have additional consumer rights 
under your local laws which this license cannot change. To the extent permitted under your 
local laws, the contributors exclude the implied warranties of merchantability, fitness for 
a particular purpose and non-infringement.

********************************************************************************************/

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsTeXProject.VisualStudio.Project
{

    #region structures

    [StructLayout(LayoutKind.Sequential)]
    internal struct _DROPFILES
    {
        public int pFiles;
        public int X;
        public int Y;
        public int fNC;
        public int fWide;
    }

    #endregion

    #region enums

    /// <summary>
    ///     The type of build performed.
    /// </summary>
    public enum BuildKind
    {
        Sync,
        Async
    }

    /// <summary>
    ///     An enumeration that describes the type of action to be taken by the build.
    /// </summary>
    [PropertyPageTypeConverter(typeof (BuildActionConverter))]
    public enum BuildAction
    {
        Compile,
        Picture,
        Content
    }

    /// <summary>
    ///     Defines the currect state of a property page.
    /// </summary>
    [Flags]
    public enum PropPageStatus
    {
        Dirty = 0x1,

        Validate = 0x2,

        Clean = 0x4
    }

    [Flags]
    [SuppressMessage("Microsoft.Design", "CA1008:EnumsShouldHaveZeroValue")]
    public enum ModuleKindFlags
    {
        ConsoleApplication,

        WindowsApplication,

        DynamicallyLinkedLibrary,

        ManifestResourceFile,

        UnmanagedDynamicallyLinkedLibrary
    }

    /// <summary>
    ///     Defines the status of the command being queried
    /// </summary>
    [Flags]
    [SuppressMessage("Microsoft.Naming", "CA1714:FlagsEnumsShouldHavePluralNames")]
    [SuppressMessage("Microsoft.Design", "CA1008:EnumsShouldHaveZeroValue")]
    public enum QueryStatusResult
    {
        /// <summary>
        ///     The command is not supported.
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "NOTSUPPORTED")] NOTSUPPORTED = 0,

        /// <summary>
        ///     The command is supported
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "SUPPORTED")] SUPPORTED = 1,

        /// <summary>
        ///     The command is enabled
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "ENABLED")] ENABLED
        = 2,

        /// <summary>
        ///     The command is toggled on
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "LATCHED")] LATCHED
        = 4,

        /// <summary>
        ///     The command is toggled off (the opposite of LATCHED).
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "NINCHED")] NINCHED
        = 8,

        /// <summary>
        ///     The command is invisible.
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "INVISIBLE")] INVISIBLE = 16
    }

    /// <summary>
    ///     Defines the type of item to be added to the hierarchy.
    /// </summary>
    public enum HierarchyAddType
    {
        AddNewItem,
        AddExistingItem
    }

    /// <summary>
    ///     Defines the component from which a command was issued.
    /// </summary>
    public enum CommandOrigin
    {
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Ui")] UiHierarchy,
        OleCommandTarget
    }

    /// <summary>
    ///     Defines the current status of the build process.
    /// </summary>
    public enum MSBuildResult
    {
        /// <summary>
        ///     The build is currently suspended.
        /// </summary>
        Suspended,

        /// <summary>
        ///     The build has been restarted.
        /// </summary>
        Resumed,

        /// <summary>
        ///     The build failed.
        /// </summary>
        Failed,

        /// <summary>
        ///     The build was successful.
        /// </summary>
        Successful
    }

    /// <summary>
    ///     Defines the type of action to be taken in showing the window frame.
    /// </summary>
    public enum WindowFrameShowAction
    {
        DoNotShow,
        Show,
        ShowNoActivate,
        Hide
    }

    /// <summary>
    ///     Defines drop types
    /// </summary>
    internal enum DropDataType
    {
        None,
        Shell,
        VsStg,
        VsRef
    }

    /// <summary>
    ///     Used by the hierarchy node to decide which element to redraw.
    /// </summary>
    [Flags]
    [SuppressMessage("Microsoft.Naming", "CA1714:FlagsEnumsShouldHavePluralNames")]
    public enum UIHierarchyElement
    {
        None = 0,

        /// <summary>
        ///     This will be translated to VSHPROPID_IconIndex
        /// </summary>
        Icon = 1,

        /// <summary>
        ///     This will be translated to VSHPROPID_StateIconIndex
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Scc")] SccState
        = 2,

        /// <summary>
        ///     This will be translated to VSHPROPID_Caption
        /// </summary>
        Caption = 4
    }

    /// <summary>
    ///     Defines the global propeties used by the msbuild project.
    /// </summary>
    public enum GlobalProperty
    {
        /// <summary>
        ///     Property specifying that we are building inside VS.
        /// </summary>
        BuildingInsideVisualStudio,

        /// <summary>
        ///     The VS installation directory. This is the same as the $(DevEnvDir) macro.
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Env")] DevEnvDir,

        /// <summary>
        ///     The name of the solution the project is created. This is the same as the $(SolutionName) macro.
        /// </summary>
        SolutionName,

        /// <summary>
        ///     The file name of the solution. This is the same as $(SolutionFileName) macro.
        /// </summary>
        SolutionFileName,

        /// <summary>
        ///     The full path of the solution. This is the same as the $(SolutionPath) macro.
        /// </summary>
        SolutionPath,

        /// <summary>
        ///     The directory of the solution. This is the same as the $(SolutionDir) macro.
        /// </summary>
        SolutionDir,

        /// <summary>
        ///     The extension of teh directory. This is the same as the $(SolutionExt) macro.
        /// </summary>
        SolutionExt,

        /// <summary>
        ///     The fxcop installation directory.
        /// </summary>
        FxCopDir,

        /// <summary>
        ///     The ResolvedNonMSBuildProjectOutputs msbuild property
        /// </summary>
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "VSIDE")] VSIDEResolvedNonMSBuildProjectOutputs,

        /// <summary>
        ///     The Configuartion property.
        /// </summary>
        Configuration,

        /// <summary>
        ///     The platform property.
        /// </summary>
        Platform,

        /// <summary>
        ///     The RunCodeAnalysisOnce property
        /// </summary>
        RunCodeAnalysisOnce,

        /// <summary>
        ///     The VisualStudioStyleErrors property
        /// </summary>
        VisualStudioStyleErrors
    }

    #endregion

    public class AfterProjectFileOpenedEventArgs : EventArgs
    {
        #region fields

        #endregion

        #region ctor

        internal AfterProjectFileOpenedEventArgs(bool added)
        {
            Added = added;
        }

        #endregion

        #region properties

        /// <summary>
        ///     True if the project is added to the solution after the solution is opened. false if the project is added to the
        ///     solution while the solution is being opened.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal bool Added { get; }

        #endregion
    }

    public class BeforeProjectFileClosedEventArgs : EventArgs
    {
        #region fields

        #endregion

        #region ctor

        internal BeforeProjectFileClosedEventArgs(bool removed)
        {
            Removed = removed;
        }

        #endregion

        #region properties

        /// <summary>
        ///     true if the project was removed from the solution before the solution was closed. false if the project was removed
        ///     from the solution while the solution was being closed.
        /// </summary>
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal bool Removed { get; }

        #endregion
    }

    /// <summary>
    ///     This class is used for the events raised by a HierarchyNode object.
    /// </summary>
    internal class HierarchyNodeEventArgs : EventArgs
    {
        internal HierarchyNodeEventArgs(HierarchyNode child)
        {
            Child = child;
        }

        public HierarchyNode Child { get; }
    }

    /// <summary>
    ///     Event args class for triggering file change event arguments.
    /// </summary>
    internal class FileChangedOnDiskEventArgs : EventArgs
    {
        /// <summary>
        ///     Constructs a new event args.
        /// </summary>
        /// <param name="fileName">File name that was changed on disk.</param>
        /// <param name="id">The item id of the file that was changed on disk.</param>
        internal FileChangedOnDiskEventArgs(string fileName, uint id, _VSFILECHANGEFLAGS flag)
        {
            this.fileName = fileName;
            itemID = id;
            fileChangeFlag = flag;
        }

        /// <summary>
        ///     Gets the file name that was changed on disk.
        /// </summary>
        /// <value>The file that was changed on disk.</value>
        internal string FileName
        {
            get { return fileName; }
        }

        /// <summary>
        ///     Gets item id of the file that has changed
        /// </summary>
        /// <value>The file that was changed on disk.</value>
        internal uint ItemID
        {
            get { return itemID; }
        }

        /// <summary>
        ///     The reason while the file has chnaged on disk.
        /// </summary>
        /// <value>The reason while the file has chnaged on disk.</value>
        internal _VSFILECHANGEFLAGS FileChangeFlag
        {
            get { return fileChangeFlag; }
        }

        #region Private fields

        /// <summary>
        ///     File name that was changed on disk.
        /// </summary>
        private readonly string fileName;

        /// <summary>
        ///     The item ide of the file that has changed.
        /// </summary>
        private readonly uint itemID;

        /// <summary>
        ///     The reason the file has changed on disk.
        /// </summary>
        private readonly _VSFILECHANGEFLAGS fileChangeFlag;

        #endregion
    }

    /// <summary>
    ///     Argument of the event raised when a project property is changed.
    /// </summary>
    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public class ProjectPropertyChangedArgs : EventArgs
    {
        internal ProjectPropertyChangedArgs(string propertyName, string oldValue, string newValue)
        {
            PropertyName = propertyName;
            OldValue = oldValue;
            NewValue = newValue;
        }

        public string NewValue { get; }

        public string OldValue { get; }

        public string PropertyName { get; }
    }
}