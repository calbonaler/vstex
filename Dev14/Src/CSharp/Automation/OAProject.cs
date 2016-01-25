﻿/********************************************************************************************

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
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsTeXProject.VisualStudio.Project.Automation
{
    [ComVisible(true), CLSCompliant(false)]
    public class OAProject : EnvDTE.Project, ISupportVSProperties
    {
        #region ctor

        public OAProject(ProjectNode project)
        {
            Project = project;
        }

        #endregion

        #region properties

        public ProjectNode Project { get; }

        #endregion

        #region ISupportVSProperties methods

        /// <summary>
        ///     Microsoft Internal Use Only.
        /// </summary>
        public virtual void NotifyPropertiesDelete()
        {
        }

        #endregion

        #region private methods

        /// <summary>
        ///     Saves or Save Asthe project.
        /// </summary>
        /// <param name="isCalledFromSaveAs">Flag determining which Save method called , the SaveAs or the Save.</param>
        /// <param name="fileName">The name of the project file.</param>
        private void DoSave(bool isCalledFromSaveAs, string fileName)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException("fileName");
            }

            if (Project == null || Project.Site == null || Project.IsClosed)
            {
                throw new InvalidOperationException();
            }

            UIThread.DoOnUIThread(delegate
            {
                using (var scope = new AutomationScope(Project.Site))
                {
                    // If an empty file name is passed in for Save then make the file name the project name.
                    if (!isCalledFromSaveAs && string.IsNullOrEmpty(fileName))
                    {
                        // Use the solution service to save the project file. Note that we have to use the service
                        // so that all the shell's elements are aware that we are inside a save operation and
                        // all the file change listenters registered by the shell are suspended.

                        // Get the cookie of the project file from the RTD.
                        var rdt = Project.Site.GetService(typeof (SVsRunningDocumentTable)) as IVsRunningDocumentTable;
                        if (null == rdt)
                        {
                            throw new InvalidOperationException();
                        }

                        IVsHierarchy hier;
                        uint itemid;
                        IntPtr unkData;
                        uint cookie;
                        ErrorHandler.ThrowOnFailure(rdt.FindAndLockDocument((uint) _VSRDTFLAGS.RDT_NoLock, Project.Url,
                            out hier,
                            out itemid, out unkData, out cookie));
                        if (IntPtr.Zero != unkData)
                        {
                            Marshal.Release(unkData);
                        }

                        // Verify that we have a cookie.
                        if (0 == cookie)
                        {
                            // This should never happen because if the project is open, then it must be in the RDT.
                            throw new InvalidOperationException();
                        }

                        // Get the IVsHierarchy for the project.
                        var prjHierarchy = Project.InteropSafeIVsHierarchy;

                        // Now get the solution.
                        var solution = Project.Site.GetService(typeof (SVsSolution)) as IVsSolution;
                        // Verify that we have both solution and hierarchy.
                        if ((null == prjHierarchy) || (null == solution))
                        {
                            throw new InvalidOperationException();
                        }

                        ErrorHandler.ThrowOnFailure(
                            solution.SaveSolutionElement((uint) __VSSLNSAVEOPTIONS.SLNSAVEOPT_SaveIfDirty, prjHierarchy,
                                cookie));
                    }
                    else
                    {
                        // We need to make some checks before we can call the save method on the project node.
                        // This is mainly because it is now us and not the caller like in  case of SaveAs or Save that should validate the file name.
                        // The IPersistFileFormat.Save method only does a validation that is necesseray to be performed. Example: in case of Save As the  
                        // file name itself is not validated only the whole path. (thus a file name like file\file is accepted, since as a path is valid)

                        // 1. The file name has to be valid. 
                        var fullPath = fileName;
                        try
                        {
                            if (!Path.IsPathRooted(fileName))
                            {
                                fullPath = Path.Combine(Project.ProjectFolder, fileName);
                            }
                        }
                            // We want to be consistent in the error message and exception we throw. fileName could be for example #�&%"�&"%  and that would trigger an ArgumentException on Path.IsRooted.
                        catch (ArgumentException)
                        {
                            throw new InvalidOperationException(SR.GetString(SR.ErrorInvalidFileName,
                                CultureInfo.CurrentUICulture));
                        }

                        // It might be redundant but we validate the file and the full path of the file being valid. The SaveAs would also validate the path.
                        // If we decide that this is performance critical then this should be refactored.
                        Utilities.ValidateFileName(Project.Site, fullPath);

                        if (!isCalledFromSaveAs)
                        {
                            // 2. The file name has to be the same 
                            if (!NativeMethods.IsSamePath(fullPath, Project.Url))
                            {
                                throw new InvalidOperationException();
                            }

                            ErrorHandler.ThrowOnFailure(Project.Save(fullPath, 1, 0));
                        }
                        else
                        {
                            ErrorHandler.ThrowOnFailure(Project.Save(fullPath, 0, 0));
                        }
                    }
                }
            });
        }

        #endregion

        #region fields

        private ConfigurationManager configurationManager;

        #endregion

        #region EnvDTE.Project

        /// <summary>
        ///     Gets or sets the name of the object.
        /// </summary>
        public virtual string Name
        {
            get { return Project.Caption; }
            set
            {
                if (Project == null || Project.Site == null || Project.IsClosed)
                {
                    throw new InvalidOperationException();
                }

                UIThread.DoOnUIThread(delegate
                {
                    using (var scope = new AutomationScope(Project.Site))
                    {
                        Project.SetEditLabel(value);
                    }
                });
            }
        }

        /// <summary>
        ///     Microsoft Internal Use Only.  Gets the file name of the project.
        /// </summary>
        public virtual string FileName
        {
            get { return Project.ProjectFile; }
        }

        /// <summary>
        ///     Microsoft Internal Use Only. Specfies if the project is dirty.
        /// </summary>
        public virtual bool IsDirty
        {
            get
            {
                int dirty;

                ErrorHandler.ThrowOnFailure(Project.IsDirty(out dirty));
                return dirty != 0;
            }
            set
            {
                if (Project == null || Project.Site == null || Project.IsClosed)
                {
                    throw new InvalidOperationException();
                }

                UIThread.DoOnUIThread(delegate
                {
                    using (var scope = new AutomationScope(Project.Site))
                    {
                        Project.SetProjectFileDirty(value);
                    }
                });
            }
        }

        /// <summary>
        ///     Gets the Projects collection containing the Project object supporting this property.
        /// </summary>
        public virtual Projects Collection
        {
            get { return null; }
        }

        /// <summary>
        ///     Gets the top-level extensibility object.
        /// </summary>
        public virtual DTE DTE
        {
            get { return (DTE) Project.Site.GetService(typeof (DTE)); }
        }

        /// <summary>
        ///     Gets a GUID string indicating the kind or type of the object.
        /// </summary>
        public virtual string Kind
        {
            get { return Project.ProjectGuid.ToString("B"); }
        }

        /// <summary>
        ///     Gets a ProjectItems collection for the Project object.
        /// </summary>
        public virtual ProjectItems ProjectItems
        {
            get { return new OAProjectItems(this, Project); }
        }

        /// <summary>
        ///     Gets a collection of all properties that pertain to the Project object.
        /// </summary>
        public virtual Properties Properties
        {
            get { return new OAProperties(Project.NodeProperties); }
        }

        /// <summary>
        ///     Returns the name of project as a relative path from the directory containing the solution file to the project file
        /// </summary>
        /// <value>Unique name if project is in a valid state. Otherwise null</value>
        public virtual string UniqueName
        {
            get
            {
                if (Project == null || Project.IsClosed)
                {
                    return null;
                }
                return UIThread.DoOnUIThread(delegate
                {
                    // Get Solution service
                    var solution = Project.GetService(typeof (IVsSolution)) as IVsSolution;
                    if (solution == null)
                    {
                        throw new InvalidOperationException();
                    }

                    // Ask solution for unique name of project
                    var uniqueName = string.Empty;
                    ErrorHandler.ThrowOnFailure(solution.GetUniqueNameOfProject(Project.InteropSafeIVsHierarchy,
                        out uniqueName));
                    return uniqueName;
                });
            }
        }

        /// <summary>
        ///     Gets an interface or object that can be accessed by name at run time.
        /// </summary>
        public virtual object Object
        {
            get { return Project.Object; }
        }

        /// <summary>
        ///     Gets the requested Extender object if it is available for this object.
        /// </summary>
        /// <param name="name">The name of the extender object.</param>
        /// <returns>An Extender object. </returns>
        public virtual object get_Extender(string name)
        {
            return null;
        }

        /// <summary>
        ///     Gets a list of available Extenders for the object.
        /// </summary>
        public virtual object ExtenderNames
        {
            get { return null; }
        }

        /// <summary>
        ///     Gets the Extender category ID (CATID) for the object.
        /// </summary>
        public virtual string ExtenderCATID
        {
            get { return string.Empty; }
        }

        /// <summary>
        ///     Gets the full path and name of the Project object's file.
        /// </summary>
        public virtual string FullName
        {
            get
            {
                return UIThread.DoOnUIThread(delegate
                {
                    string filename;
                    uint format;
                    ErrorHandler.ThrowOnFailure(Project.GetCurFile(out filename, out format));
                    return filename;
                });
            }
        }

        /// <summary>
        ///     Gets or sets a value indicatingwhether the object has not been modified since last being saved or opened.
        /// </summary>
        public virtual bool Saved
        {
            get { return !IsDirty; }
            set
            {
                if (Project == null || Project.Site == null || Project.IsClosed)
                {
                    throw new InvalidOperationException();
                }

                UIThread.DoOnUIThread(delegate
                {
                    using (var scope = new AutomationScope(Project.Site))
                    {
                        Project.SetProjectFileDirty(!value);
                    }
                });
            }
        }

        /// <summary>
        ///     Gets the ConfigurationManager object for this Project .
        /// </summary>
        public virtual ConfigurationManager ConfigurationManager
        {
            get
            {
                return UIThread.DoOnUIThread(delegate
                {
                    if (configurationManager == null)
                    {
                        var extensibility = Project.Site.GetService(typeof (IVsExtensibility)) as IVsExtensibility3;

                        if (extensibility == null)
                        {
                            throw new InvalidOperationException();
                        }

                        object configurationManagerAsObject;
                        ErrorHandler.ThrowOnFailure(extensibility.GetConfigMgr(Project.InteropSafeIVsHierarchy,
                            VSConstants.VSITEMID_ROOT, out configurationManagerAsObject));

                        if (configurationManagerAsObject == null)
                        {
                            throw new InvalidOperationException();
                        }
                        configurationManager = (ConfigurationManager) configurationManagerAsObject;
                    }

                    return configurationManager;
                });
            }
        }

        /// <summary>
        ///     Gets the Globals object containing add-in values that may be saved in the solution (.sln) file, the project file,
        ///     or in the user's profile data.
        /// </summary>
        public virtual Globals Globals
        {
            get { return null; }
        }

        /// <summary>
        ///     Gets a ProjectItem object for the nested project in the host project.
        /// </summary>
        public virtual ProjectItem ParentProjectItem
        {
            get { return null; }
        }

        /// <summary>
        ///     Gets the CodeModel object for the project.
        /// </summary>
        public virtual CodeModel CodeModel
        {
            get { return null; }
        }

        /// <summary>
        ///     Saves the project.
        /// </summary>
        /// <param name="fileName">
        ///     The file name with which to save the solution, project, or project item. If the file exists, it
        ///     is overwritten
        /// </param>
        /// <exception cref="InvalidOperationException">Is thrown if the save operation failes.</exception>
        /// <exception cref="ArgumentNullException">Is thrown if fileName is null.</exception>
        public virtual void SaveAs(string fileName)
        {
            DoSave(true, fileName);
        }

        /// <summary>
        ///     Saves the project
        /// </summary>
        /// <param name="fileName">The file name of the project</param>
        /// <exception cref="InvalidOperationException">Is thrown if the save operation failes.</exception>
        /// <exception cref="ArgumentNullException">Is thrown if fileName is null.</exception>
        public virtual void Save(string fileName)
        {
            DoSave(false, fileName);
        }

        /// <summary>
        ///     Removes the project from the current solution.
        /// </summary>
        public virtual void Delete()
        {
            if (Project == null || Project.Site == null || Project.IsClosed)
            {
                throw new InvalidOperationException();
            }

            UIThread.DoOnUIThread(delegate
            {
                using (var scope = new AutomationScope(Project.Site))
                {
                    Project.Remove(false);
                }
            });
        }

        #endregion
    }
}