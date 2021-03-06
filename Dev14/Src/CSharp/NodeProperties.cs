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
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace VsTeXProject.VisualStudio.Project
{
    /// <summary>
    ///     All public properties on Nodeproperties or derived classes are assumed to be used by Automation by default.
    ///     Set this attribute to false on Properties that should not be visible for Automation.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class AutomationBrowsableAttribute : Attribute
    {
        public AutomationBrowsableAttribute(bool browsable)
        {
            Browsable = browsable;
        }

        public bool Browsable { get; }
    }

    /// <summary>
    ///     To create your own localizable node properties, subclass this and add public properties
    ///     decorated with your own localized display name, category and description attributes.
    /// </summary>
    [CLSCompliant(false), ComVisible(true)]
    public class NodeProperties : LocalizableProperties,
        ISpecifyPropertyPages,
        IVsGetCfgProvider,
        IVsSpecifyProjectDesignerPages,
        IInternalExtenderProvider,
        IVsBrowseObject
    {
        #region fields

        #endregion

        #region ctors

        public NodeProperties(HierarchyNode node)
        {
            if (node == null)
            {
                throw new ArgumentNullException("node");
            }
            Node = node;
        }

        #endregion

        #region ISpecifyPropertyPages methods

        public virtual void GetPages(CAUUID[] pages)
        {
            GetCommonPropertyPages(pages);
        }

        #endregion

        #region IVsBrowseObject methods

        /// <summary>
        ///     Maps back to the hierarchy or project item object corresponding to the browse object.
        /// </summary>
        /// <param name="hier">Reference to the hierarchy object.</param>
        /// <param name="itemid">Reference to the project item.</param>
        /// <returns>If the method succeeds, it returns S_OK. If it fails, it returns an error code. </returns>
        public virtual int GetProjectItem(out IVsHierarchy hier, out uint itemid)
        {
            if (Node == null)
            {
                throw new InvalidOperationException();
            }
            hier = Node.ProjectMgr.InteropSafeIVsHierarchy;
            itemid = Node.ID;
            return VSConstants.S_OK;
        }

        #endregion

        #region IVsGetCfgProvider methods

        public virtual int GetCfgProvider(out IVsCfgProvider p)
        {
            p = null;
            return VSConstants.E_NOTIMPL;
        }

        #endregion

        #region IVsSpecifyProjectDesignerPages

        /// <summary>
        ///     Implementation of the IVsSpecifyProjectDesignerPages. It will retun the pages that are configuration independent.
        /// </summary>
        /// <param name="pages">The pages to return.</param>
        /// <returns></returns>
        public virtual int GetProjectDesignerPages(CAUUID[] pages)
        {
            GetCommonPropertyPages(pages);
            return VSConstants.S_OK;
        }

        #endregion

        #region overridden methods

        /// <summary>
        ///     Get the Caption of the Hierarchy Node instance. If Caption is null or empty we delegate to base
        /// </summary>
        /// <returns>Caption of Hierarchy node instance</returns>
        public override string GetComponentName()
        {
            var caption = Node.Caption;
            if (string.IsNullOrEmpty(caption))
            {
                return base.GetComponentName();
            }
            return caption;
        }

        #endregion

        #region properties

        [Browsable(false)]
        [AutomationBrowsable(false)]
        public HierarchyNode Node { get; }

        /// <summary>
        ///     Used by Property Pages Frame to set it's title bar. The Caption of the Hierarchy Node is returned.
        /// </summary>
        [Browsable(false)]
        [AutomationBrowsable(false)]
        public virtual string Name
        {
            get { return Node.Caption; }
        }

        #endregion

        #region helper methods

        protected string GetProperty(string name, string def)
        {
            var a = Node.ItemNode.GetMetadata(name);
            return a == null ? def : a;
        }

        protected void SetProperty(string name, string value)
        {
            Node.ItemNode.SetMetadata(name, value);
        }

        /// <summary>
        ///     Retrieves the common property pages. The NodeProperties is the BrowseObject and that will be called to support
        ///     configuration independent properties.
        /// </summary>
        /// <param name="pages">The pages to return.</param>
        private void GetCommonPropertyPages(CAUUID[] pages)
        {
            // We do not check whether the supportsProjectDesigner is set to false on the ProjectNode.
            // We rely that the caller knows what to call on us.
            if (pages == null)
            {
                throw new ArgumentNullException("pages");
            }

            if (pages.Length == 0)
            {
                throw new ArgumentException(SR.GetString(SR.InvalidParameter, CultureInfo.CurrentUICulture), "pages");
            }

            // Only the project should show the property page the rest should show the project properties.
            if (Node != null && Node is ProjectNode)
            {
                // Retrieve the list of guids from hierarchy properties.
                // Because a flavor could modify that list we must make sure we are calling the outer most implementation of IVsHierarchy
                var guidsList = string.Empty;
                var hierarchy = Node.ProjectMgr.InteropSafeIVsHierarchy;
                object variant = null;
                ErrorHandler.ThrowOnFailure(hierarchy.GetProperty(VSConstants.VSITEMID_ROOT,
                    (int) __VSHPROPID2.VSHPROPID_PropertyPagesCLSIDList, out variant));
                guidsList = (string) variant;

                var guids = Utilities.GuidsArrayFromSemicolonDelimitedStringOfGuids(guidsList);
                if (guids == null || guids.Length == 0)
                {
                    pages[0] = new CAUUID();
                    pages[0].cElems = 0;
                }
                else
                {
                    pages[0] = PackageUtilities.CreateCAUUIDFromGuidArray(guids);
                }
            }
            else
            {
                pages[0] = new CAUUID();
                pages[0].cElems = 0;
            }
        }

        #endregion

        #region IInternalExtenderProvider Members

        bool IInternalExtenderProvider.CanExtend(string extenderCATID, string extenderName, object extendeeObject)
        {
            var outerHierarchy = Node.ProjectMgr.InteropSafeIVsHierarchy as IInternalExtenderProvider;


            if (outerHierarchy != null)
            {
                return outerHierarchy.CanExtend(extenderCATID, extenderName, extendeeObject);
            }
            return false;
        }

        object IInternalExtenderProvider.GetExtender(string extenderCATID, string extenderName, object extendeeObject,
            IExtenderSite extenderSite, int cookie)
        {
            var outerHierarchy = Node.ProjectMgr.InteropSafeIVsHierarchy as IInternalExtenderProvider;

            if (outerHierarchy != null)
            {
                return outerHierarchy.GetExtender(extenderCATID, extenderName, extendeeObject, extenderSite, cookie);
            }

            return null;
        }

        object IInternalExtenderProvider.GetExtenderNames(string extenderCATID, object extendeeObject)
        {
            var outerHierarchy = Node.ProjectMgr.InteropSafeIVsHierarchy as IInternalExtenderProvider;

            if (outerHierarchy != null)
            {
                return outerHierarchy.GetExtenderNames(extenderCATID, extendeeObject);
            }

            return null;
        }

        #endregion

        #region ExtenderSupport

        [Browsable(false)]
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "CATID")]
        public virtual string ExtenderCATID
        {
            get
            {
                var catid = Node.ProjectMgr.GetCATIDForType(GetType());
                if (Guid.Empty.CompareTo(catid) == 0)
                {
                    return null;
                }
                return catid.ToString("B");
            }
        }

        [Browsable(false)]
        public object ExtenderNames()
        {
            var extenderService = (ObjectExtenders) Node.GetService(typeof (ObjectExtenders));
            Debug.Assert(extenderService != null,
                "Could not get the ObjectExtenders object from the services exposed by this property object");
            if (extenderService == null)
            {
                throw new InvalidOperationException();
            }
            return extenderService.GetExtenderNames(ExtenderCATID, this);
        }

        public object Extender(string extenderName)
        {
            var extenderService = (ObjectExtenders) Node.GetService(typeof (ObjectExtenders));
            Debug.Assert(extenderService != null,
                "Could not get the ObjectExtenders object from the services exposed by this property object");
            if (extenderService == null)
            {
                throw new InvalidOperationException();
            }
            return extenderService.GetExtender(ExtenderCATID, extenderName, this);
        }

        #endregion
    }

    [CLSCompliant(false), ComVisible(true)]
    public class FileNodeProperties : NodeProperties
    {
        #region ctors

        public FileNodeProperties(HierarchyNode node)
            : base(node)
        {
        }

        #endregion

        #region overridden methods

        public override string GetClassName()
        {
            return SR.GetString(SR.FileProperties, CultureInfo.CurrentUICulture);
        }

        #endregion

        #region properties

        [SRCategory(SR.Advanced)]
        [LocDisplayName(SR.BuildAction)]
        [SRDescription(SR.BuildActionDescription)]
        public virtual BuildAction BuildAction
        {
            get
            {
                var value = Node.ItemNode.ItemName;
                if (value == null || value.Length == 0)
                {
                    return BuildAction.Content;
                }
                return (BuildAction) Enum.Parse(typeof (BuildAction), value);
            }
            set { Node.ItemNode.ItemName = value.ToString(); }
        }

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.FileName)]
        [SRDescription(SR.FileNameDescription)]
        public string FileName
        {
            get { return Node.Caption; }
            set { Node.SetEditLabel(value); }
        }

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.FullPath)]
        [SRDescription(SR.FullPathDescription)]
        public string FullPath
        {
            get { return Node.Url; }
        }

        #region non-browsable properties - used for automation only

        [Browsable(false)]
        public string Extension
        {
            get { return Path.GetExtension(Node.Caption); }
        }

        #endregion

        #endregion
    }

    [CLSCompliant(false), ComVisible(true)]
    public class DependentFileNodeProperties : NodeProperties
    {
        #region ctors

        public DependentFileNodeProperties(HierarchyNode node)
            : base(node)
        {
        }

        #endregion

        #region overridden methods

        public override string GetClassName()
        {
            return SR.GetString(SR.FileProperties, CultureInfo.CurrentUICulture);
        }

        #endregion

        #region properties

        [SRCategory(SR.Advanced)]
        [LocDisplayName(SR.BuildAction)]
        [SRDescription(SR.BuildActionDescription)]
        public virtual BuildAction BuildAction
        {
            get
            {
                var value = Node.ItemNode.ItemName;
                if (value == null || value.Length == 0)
                {
                    return BuildAction.Content;
                }
                return (BuildAction) Enum.Parse(typeof (BuildAction), value);
            }
            set { Node.ItemNode.ItemName = value.ToString(); }
        }

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.FileName)]
        [SRDescription(SR.FileNameDescription)]
        public virtual string FileName
        {
            get { return Node.Caption; }
        }

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.FullPath)]
        [SRDescription(SR.FullPathDescription)]
        public string FullPath
        {
            get { return Node.Url; }
        }

        #endregion
    }

    [CLSCompliant(false), ComVisible(true)]
    public class SingleFileGeneratorNodeProperties : FileNodeProperties
    {
        #region fields

        private EventHandler<HierarchyNodeEventArgs> onCustomToolChanged;

        #endregion

        #region ctors

        public SingleFileGeneratorNodeProperties(HierarchyNode node)
            : base(node)
        {

        }

        #endregion

        #region properties

        [SRCategory(SR.Advanced)]
        [LocDisplayName(SR.CustomTool)]
        [SRDescription(SR.CustomToolDescription)]
        public virtual string CustomTool
        {
            get { return Node.ItemNode.GetMetadata(ProjectFileConstants.Generator); }
            set
            {
                if (CustomTool != value)
                {
                    Node.ItemNode.SetMetadata(ProjectFileConstants.Generator, !string.IsNullOrEmpty(value) ? value : null);
					onCustomToolChanged?.Invoke(Node, new HierarchyNodeEventArgs(Node));
				}
            }
        }

        #endregion

        #region custom tool events

        internal event EventHandler<HierarchyNodeEventArgs> OnCustomToolChanged
        {
            add { onCustomToolChanged += value; }
            remove { onCustomToolChanged -= value; }
        }

        #endregion
    }

    [CLSCompliant(false), ComVisible(true)]
    public class ProjectNodeProperties : NodeProperties
    {
        #region ctors

        public ProjectNodeProperties(ProjectNode node)
            : base(node)
        {
        }

        #endregion

        #region properties

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.ProjectFolder)]
        [SRDescription(SR.ProjectFolderDescription)]
        [AutomationBrowsable(false)]
        public string ProjectFolder
        {
            get { return Node.ProjectMgr.ProjectFolder; }
        }

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.ProjectFile)]
        [SRDescription(SR.ProjectFileDescription)]
        [AutomationBrowsable(false)]
        public string ProjectFile
        {
            get { return Node.ProjectMgr.ProjectFile; }
            set { Node.ProjectMgr.ProjectFile = value; }
        }

        #region non-browsable properties - used for automation only

        [Browsable(false)]
        public string FileName
        {
            get { return Node.ProjectMgr.ProjectFile; }
            set { Node.ProjectMgr.ProjectFile = value; }
        }


        [Browsable(false)]
        public string FullPath
        {
            get
            {
                var fullPath = Node.ProjectMgr.ProjectFolder;
                if (!fullPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                {
                    return fullPath + Path.DirectorySeparatorChar;
                }
                return fullPath;
            }
        }

        #endregion

        #endregion

        #region overridden methods

        public override string GetClassName()
        {
            return SR.GetString(SR.ProjectProperties, CultureInfo.CurrentUICulture);
        }

        /// <summary>
        ///     ICustomTypeDescriptor.GetEditor
        ///     To enable the "Property Pages" button on the properties browser
        ///     the browse object (project properties) need to be unmanaged
        ///     or it needs to provide an editor of type ComponentEditor.
        /// </summary>
        /// <param name="editorBaseType">Type of the editor</param>
        /// <returns>Editor</returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "The service provider is used by the PropertiesEditorLauncher")]
        public override object GetEditor(Type editorBaseType)
        {
            // Override the scenario where we are asked for a ComponentEditor
            // as this is how the Properties Browser calls us
            if (editorBaseType == typeof (ComponentEditor))
            {
                IOleServiceProvider sp;
                ErrorHandler.ThrowOnFailure(Node.GetSite(out sp));
                return new PropertiesEditorLauncher(new ServiceProvider(sp));
            }

            return base.GetEditor(editorBaseType);
        }

        public override int GetCfgProvider(out IVsCfgProvider p)
        {
            if (Node != null && Node.ProjectMgr != null)
            {
                return Node.ProjectMgr.GetCfgProvider(out p);
            }

            return base.GetCfgProvider(out p);
        }

        #endregion
    }

    [CLSCompliant(false), ComVisible(true)]
    public class FolderNodeProperties : NodeProperties
    {
        #region ctors

        public FolderNodeProperties(HierarchyNode node)
            : base(node)
        {
        }

        #endregion

        #region overridden methods

        public override string GetClassName()
        {
            return SR.GetString(SR.FolderProperties, CultureInfo.CurrentUICulture);
        }

        #endregion

        #region properties

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.FolderName)]
        [SRDescription(SR.FolderNameDescription)]
        [AutomationBrowsable(false)]
        public string FolderName
        {
            get { return Node.Caption; }
            set
            {
                Node.SetEditLabel(value);
                Node.ReDraw(UIHierarchyElement.Caption);
            }
        }

        #region properties - used for automation only

        [Browsable(false)]
        [AutomationBrowsable(true)]
        public string FileName
        {
            get { return Node.Caption; }
            set { Node.SetEditLabel(value); }
        }

        [Browsable(false)]
        [AutomationBrowsable(true)]
        public string FullPath
        {
            get
            {
                var fullPath = Node.GetMkDocument();
                if (!fullPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                {
                    return fullPath + Path.DirectorySeparatorChar;
                }
                return fullPath;
            }
        }

        #endregion

        #endregion
    }

    [CLSCompliant(false), ComVisible(true)]
    public class ReferenceNodeProperties : NodeProperties
    {
        #region ctors

        public ReferenceNodeProperties(HierarchyNode node)
            : base(node)
        {
        }

        #endregion

        #region overridden methods

        public override string GetClassName()
        {
            return SR.GetString(SR.ReferenceProperties, CultureInfo.CurrentUICulture);
        }

        #endregion

        #region properties

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.RefName)]
        [SRDescription(SR.RefNameDescription)]
        [Browsable(true)]
        [AutomationBrowsable(true)]
        public override string Name
        {
            get { return Node.Caption; }
        }

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.CopyToLocal)]
        [SRDescription(SR.CopyToLocalDescription)]
        public bool CopyToLocal
        {
            get
            {
                var copyLocal = GetProperty(ProjectFileConstants.Private, "False");
                if (copyLocal == null || copyLocal.Length == 0)
                    return true;
                return bool.Parse(copyLocal);
            }
            set { SetProperty(ProjectFileConstants.Private, value.ToString()); }
        }

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.FullPath)]
        [SRDescription(SR.FullPathDescription)]
        public virtual string FullPath
        {
            get { return Node.Url; }
        }

        #endregion
    }

    [ComVisible(true)]
    public class ProjectReferencesProperties : ReferenceNodeProperties
    {
        #region ctors

        public ProjectReferencesProperties(ProjectReferenceNode node)
            : base(node)
        {
        }

        #endregion

        #region overriden methods

        public override string FullPath
        {
            get { return ((ProjectReferenceNode) Node).ReferencedProjectOutputPath; }
        }

        #endregion
    }

    [ComVisible(true)]
    public class ComReferenceProperties : ReferenceNodeProperties
    {
        public ComReferenceProperties(ComReferenceNode node)
            : base(node)
        {
        }

        [SRCategory(SR.Misc)]
        [LocDisplayName(SR.EmbedInteropTypes)]
        [SRDescription(SR.EmbedInteropTypesDescription)]
        public virtual bool EmbedInteropTypes
        {
            get { return ((ComReferenceNode) Node).EmbedInteropTypes; }
            set { ((ComReferenceNode) Node).EmbedInteropTypes = value; }
        }
    }
}