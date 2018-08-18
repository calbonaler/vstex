using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using VsTeXProject.VisualStudio.Project;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsTeXProject
{
    /// <summary>
    /// This class implements general property page for the project type.
    /// </summary>
    [ComVisible(true)]
    [Guid("5F9F1697-2E61-4c10-9AD2-94FA2A9BAAE8")]
    public class GeneralPropertyPage : SettingsPage
    {
        #region Fields
        private string _teXProcessor;
        private string _toolspath;

        #endregion Fields

        #region Constructors
        /// <summary>
        /// Explicitly defined default constructor.
        /// </summary>
        public GeneralPropertyPage()
        {
            this.Name = Resources.GetString(Resources.GeneralCaption);
        }
        #endregion

        #region Properties

        [ResourcesCategoryAttribute(Resources.Application)]
        [LocDisplayName(Resources.TeXProcessor)]
        [ResourcesDescriptionAttribute(Resources.TeXProcessorDescription)]
		[PropertyPageTypeConverter(typeof(TeXProcessorConverter))]
        /// <summary>
        /// Gets or sets TeX Processor.
        /// </summary>
        /// <remarks>IsDirty flag was switched to true.</remarks>
        public string TeXProcessor
        {
            get { return this._teXProcessor; }
            set { this._teXProcessor = value; this.IsDirty = true; }
        }

        [ResourcesCategoryAttribute(Resources.Application)]
        [LocDisplayName(Resources.ToolsPath)]
        [ResourcesDescriptionAttribute(Resources.ToolsPathDescription)]
        /// <summary>
        /// Gets or sets Tools Path.
        /// </summary>
        /// <remarks>IsDirty flag was switched to true.</remarks>
        public string ToolsPath
        {
            get { return this._toolspath; }
            set { this._toolspath = value; this.IsDirty = true; }
        }

        [ResourcesCategoryAttribute(Resources.Project)]
        [LocDisplayName(Resources.ProjectFile)]
        [ResourcesDescriptionAttribute(Resources.ProjectFileDescription)]
        /// <summary>
        /// Gets the path to the project file.
        /// </summary>
        /// <remarks>IsDirty flag was switched to true.</remarks>
        public string ProjectFile
        {
            get { return Path.GetFileName(this.ProjectMgr.ProjectFile); }
        }

        [ResourcesCategoryAttribute(Resources.Project)]
        [LocDisplayName(Resources.ProjectFolder)]
        [ResourcesDescriptionAttribute(Resources.ProjectFolderDescription)]
        /// <summary>
        /// Gets the path to the project folder.
        /// </summary>
        /// <remarks>IsDirty flag was switched to true.</remarks>
        public string ProjectFolder
        {
            get { return Path.GetDirectoryName(this.ProjectMgr.ProjectFolder); }
        }

        #endregion

        #region Overriden Implementation
        /// <summary>
        /// Returns class FullName property value.
        /// </summary>
        public override string GetClassName()
        {
            return this.GetType().FullName;
        }

        /// <summary>
        /// Bind properties.
        /// </summary>
        protected override void BindProperties()
        {
            if(this.ProjectMgr == null)
            {
                return;
            }

            this._teXProcessor = this.ProjectMgr.GetProjectProperty("TeXProcessor", false);
            this._toolspath = this.ProjectMgr.GetProjectProperty("ToolsPath", false);
        }

        /// <summary>
        /// Apply Changes on project node.
        /// </summary>
        /// <returns>E_INVALIDARG if internal ProjectMgr is null, otherwise applies changes and return S_OK.</returns>
        protected override int ApplyChanges()
        {
            if(this.ProjectMgr == null)
            {
                return VSConstants.E_INVALIDARG;
            }

            this.ProjectMgr.SetProjectProperty("TeXProcessor", this.TeXProcessor);
            this.ProjectMgr.SetProjectProperty("ToolsPath",this._toolspath);

            this.IsDirty = false;

            return VSConstants.S_OK;
        }
        #endregion
    }
}
