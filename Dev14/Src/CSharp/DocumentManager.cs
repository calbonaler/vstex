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
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShellConstants = Microsoft.VisualStudio.Shell.Interop.Constants;

namespace VsTeXProject.VisualStudio.Project
{
    /// <summary>
    ///     This abstract class handles opening, saving of items in the hierarchy.
    /// </summary>
    [CLSCompliant(false)]
    public abstract class DocumentManager
    {
        #region fields

        #endregion

        #region ctors

        protected DocumentManager(HierarchyNode node)
        {
            Node = node;
        }

        #endregion

        #region properties

        protected HierarchyNode Node { get; }

        #endregion

        #region virtual methods

        /// <summary>
        ///     Open a document using the standard editor. This method has no implementation since a document is abstract in this
        ///     context
        /// </summary>
        /// <param name="logicalView">
        ///     In MultiView case determines view to be activated by IVsMultiViewDocumentView. For a list of
        ///     logical view GUIDS, see constants starting with LOGVIEWID_ defined in NativeMethods class
        /// </param>
        /// <param name="docDataExisting">IntPtr to the IUnknown interface of the existing document data object</param>
        /// <param name="windowFrame">A reference to the window frame that is mapped to the document</param>
        /// <param name="windowFrameAction">Determine the UI action on the document window</param>
        /// <returns>NotImplementedException</returns>
        /// <remarks>See FileDocumentManager class for an implementation of this method</remarks>
        public virtual int Open(ref Guid logicalView, IntPtr docDataExisting, out IVsWindowFrame windowFrame,
            WindowFrameShowAction windowFrameAction)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        ///     Open a document using a specific editor. This method has no implementation.
        /// </summary>
        /// <param name="editorFlags">
        ///     Specifies actions to take when opening a specific editor. Possible editor flags are defined
        ///     in the enumeration Microsoft.VisualStudio.Shell.Interop.__VSOSPEFLAGS
        /// </param>
        /// <param name="editorType">Unique identifier of the editor type</param>
        /// <param name="physicalView">
        ///     Name of the physical view. If null, the environment calls MapLogicalView on the editor
        ///     factory to determine the physical view that corresponds to the logical view. In this case, null does not specify
        ///     the primary view, but rather indicates that you do not know which view corresponds to the logical view
        /// </param>
        /// <param name="logicalView">
        ///     In MultiView case determines view to be activated by IVsMultiViewDocumentView. For a list of
        ///     logical view GUIDS, see constants starting with LOGVIEWID_ defined in NativeMethods class
        /// </param>
        /// <param name="docDataExisting">IntPtr to the IUnknown interface of the existing document data object</param>
        /// <param name="frame">A reference to the window frame that is mapped to the document</param>
        /// <param name="windowFrameAction">Determine the UI action on the document window</param>
        /// <returns>NotImplementedException</returns>
        /// <remarks>See FileDocumentManager for an implementation of this method</remarks>
        public virtual int OpenWithSpecific(uint editorFlags, ref Guid editorType, string physicalView,
            ref Guid logicalView, IntPtr docDataExisting, out IVsWindowFrame frame,
            WindowFrameShowAction windowFrameAction)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        ///     Close an open document window
        /// </summary>
        /// <param name="closeFlag">Decides how to close the document</param>
        /// <returns>S_OK if successful, otherwise an error is returned</returns>
        public virtual int Close(__FRAMECLOSE closeFlag)
        {
            if (Node == null || Node.ProjectMgr == null || Node.ProjectMgr.IsClosed)
            {
                return VSConstants.E_FAIL;
            }

            // Get info about the document
            bool isDirty, isOpen, isOpenedByUs;
            uint docCookie;
            IVsPersistDocData ppIVsPersistDocData;
            GetDocInfo(out isOpen, out isDirty, out isOpenedByUs, out docCookie, out ppIVsPersistDocData);

            if (isOpenedByUs)
            {
                var shell = Node.ProjectMgr.Site.GetService(typeof (IVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
                var logicalView = Guid.Empty;
                uint grfIDO = 0;
                IVsUIHierarchy pHierOpen;
                var itemIdOpen = new uint[1];
                IVsWindowFrame windowFrame;
                int fOpen;
                ErrorHandler.ThrowOnFailure(shell.IsDocumentOpen(Node.ProjectMgr.InteropSafeIVsUIHierarchy, Node.ID,
                    Node.Url, ref logicalView, grfIDO, out pHierOpen, itemIdOpen, out windowFrame, out fOpen));

                if (windowFrame != null)
                {
                    docCookie = 0;
                    return windowFrame.CloseFrame((uint) closeFlag);
                }
            }

            return VSConstants.S_OK;
        }

        /// <summary>
        ///     Silently saves an open document
        /// </summary>
        /// <param name="saveIfDirty">Save the open document only if it is dirty</param>
        /// <remarks>
        ///     The call to SaveDocData may return Microsoft.VisualStudio.Shell.Interop.PFF_RESULTS.STG_S_DATALOSS to indicate
        ///     some characters could not be represented in the current codepage
        /// </remarks>
        public virtual void Save(bool saveIfDirty)
        {
            bool isDirty, isOpen, isOpenedByUs;
            uint docCookie;
            IVsPersistDocData persistDocData;
            GetDocInfo(out isOpen, out isDirty, out isOpenedByUs, out docCookie, out persistDocData);
            if (isDirty && saveIfDirty && persistDocData != null)
            {
                string name;
                int cancelled;
                ErrorHandler.ThrowOnFailure(persistDocData.SaveDocData(VSSAVEFLAGS.VSSAVE_SilentSave, out name,
                    out cancelled));
            }
        }

        #endregion

        #region helper methods

        /// <summary>
        ///     Get document properties from RDT
        /// </summary>
        internal void GetDocInfo(
            out bool isOpen, // true if the doc is opened
            out bool isDirty, // true if the doc is dirty
            out bool isOpenedByUs, // true if opened by our project
            out uint docCookie, // VSDOCCOOKIE if open
            out IVsPersistDocData persistDocData)
        {
            isOpen = isDirty = isOpenedByUs = false;
            docCookie = (uint) ShellConstants.VSDOCCOOKIE_NIL;
            persistDocData = null;

            if (Node == null || Node.ProjectMgr == null || Node.ProjectMgr.IsClosed)
            {
                return;
            }

            IVsHierarchy hierarchy;
            var vsitemid = VSConstants.VSITEMID_NIL;

            VsShellUtilities.GetRDTDocumentInfo(Node.ProjectMgr.Site, Node.Url, out hierarchy, out vsitemid,
                out persistDocData, out docCookie);

            if (hierarchy == null || docCookie == (uint) ShellConstants.VSDOCCOOKIE_NIL)
            {
                return;
            }

            isOpen = true;
            // check if the doc is opened by another project
            if (Utilities.IsSameComObject(Node.ProjectMgr.InteropSafeIVsHierarchy, hierarchy))
            {
                isOpenedByUs = true;
            }

            if (persistDocData != null)
            {
                int isDocDataDirty;
                ErrorHandler.ThrowOnFailure(persistDocData.IsDocDataDirty(out isDocDataDirty));
                isDirty = isDocDataDirty != 0;
            }
        }

        protected string GetOwnerCaption()
        {
            Debug.Assert(Node != null, "No node has been initialized for the document manager");

            object pvar;
            ErrorHandler.ThrowOnFailure(Node.GetProperty(Node.ID, (int) __VSHPROPID.VSHPROPID_Caption, out pvar));

            return pvar as string;
        }

        protected static void CloseWindowFrame(ref IVsWindowFrame windowFrame)
        {
            if (windowFrame != null)
            {
                try
                {
                    Marshal.ThrowExceptionForHR(windowFrame.CloseFrame(0));
                }
                finally
                {
                    windowFrame = null;
                }
            }
        }

        protected string GetFullPathForDocument()
        {
            var fullPath = string.Empty;

            Debug.Assert(Node != null, "No node has been initialized for the document manager");

            // Get the URL representing the item
            fullPath = Node.GetMkDocument();

            Debug.Assert(!string.IsNullOrEmpty(fullPath),
                "Could not retrive the fullpath for the node" + Node.ID.ToString(CultureInfo.CurrentCulture));
            return fullPath;
        }

        #endregion

        #region static methods

        /// <summary>
        ///     Updates the caption for all windows associated to the document.
        /// </summary>
        /// <param name="site">The service provider.</param>
        /// <param name="caption">The new caption.</param>
        /// <param name="docData">The IUnknown interface to a document data object associated with a registered document.</param>
        public static void UpdateCaption(IServiceProvider site, string caption, IntPtr docData)
        {
            if (site == null)
            {
                throw new ArgumentNullException("site");
            }

            if (string.IsNullOrEmpty(caption))
            {
                throw new ArgumentException(
                    SR.GetString(SR.ParameterCannotBeNullOrEmpty, CultureInfo.CurrentUICulture), "caption");
            }

            var uiShell = site.GetService(typeof (SVsUIShell)) as IVsUIShell;

            // We need to tell the windows to update their captions. 
            IEnumWindowFrames windowFramesEnum;
            ErrorHandler.ThrowOnFailure(uiShell.GetDocumentWindowEnum(out windowFramesEnum));
            var windowFrames = new IVsWindowFrame[1];
            uint fetched;
            while (windowFramesEnum.Next(1, windowFrames, out fetched) == VSConstants.S_OK && fetched == 1)
            {
                var windowFrame = windowFrames[0];
                object data;
                ErrorHandler.ThrowOnFailure(windowFrame.GetProperty((int) __VSFPROPID.VSFPROPID_DocData, out data));
                var ptr = Marshal.GetIUnknownForObject(data);
                try
                {
                    if (ptr == docData)
                    {
                        ErrorHandler.ThrowOnFailure(windowFrame.SetProperty((int) __VSFPROPID.VSFPROPID_OwnerCaption,
                            caption));
                    }
                }
                finally
                {
                    if (ptr != IntPtr.Zero)
                    {
                        Marshal.Release(ptr);
                    }
                }
            }
        }

        /// <summary>
        ///     Rename document in the running document table from oldName to newName.
        /// </summary>
        /// <param name="provider">The service provider.</param>
        /// <param name="oldName">Full path to the old name of the document.</param>
        /// <param name="newName">Full path to the new name of the document.</param>
        /// <param name="newItemId">The new item id of the document</param>
        public static void RenameDocument(IServiceProvider site, string oldName, string newName, uint newItemId)
        {
            if (site == null)
            {
                throw new ArgumentNullException("site");
            }

            if (string.IsNullOrEmpty(oldName))
            {
                throw new ArgumentException(
                    SR.GetString(SR.ParameterCannotBeNullOrEmpty, CultureInfo.CurrentUICulture), "oldName");
            }

            if (string.IsNullOrEmpty(newName))
            {
                throw new ArgumentException(
                    SR.GetString(SR.ParameterCannotBeNullOrEmpty, CultureInfo.CurrentUICulture), "newName");
            }

            if (newItemId == VSConstants.VSITEMID_NIL)
            {
                throw new ArgumentNullException("newItemId");
            }

            var pRDT = site.GetService(typeof (SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            var doc = site.GetService(typeof (SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;

            if (pRDT == null || doc == null) return;

            IVsHierarchy pIVsHierarchy;
            uint itemId;
            IntPtr docData;
            uint uiVsDocCookie;
            ErrorHandler.ThrowOnFailure(pRDT.FindAndLockDocument((uint) _VSRDTFLAGS.RDT_NoLock, oldName,
                out pIVsHierarchy, out itemId, out docData, out uiVsDocCookie));

            if (docData != IntPtr.Zero)
            {
                try
                {
                    var pUnk = Marshal.GetIUnknownForObject(pIVsHierarchy);
                    var iid = typeof (IVsHierarchy).GUID;
                    IntPtr pHier;
                    Marshal.QueryInterface(pUnk, ref iid, out pHier);
                    try
                    {
                        ErrorHandler.ThrowOnFailure(pRDT.RenameDocument(oldName, newName, pHier, newItemId));
                    }
                    finally
                    {
                        if (pHier != IntPtr.Zero)
                            Marshal.Release(pHier);
                        if (pUnk != IntPtr.Zero)
                            Marshal.Release(pUnk);
                    }
                }
                finally
                {
                    Marshal.Release(docData);
                }
            }
        }

        #endregion
    }
}