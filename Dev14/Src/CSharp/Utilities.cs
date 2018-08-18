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
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using IServiceProvider = System.IServiceProvider;
using MSBuild = Microsoft.Build.Evaluation;

namespace VsTeXProject.VisualStudio.Project
{
    public static class Utilities
    {
        private static readonly char[] curlyBraces = {'{', '}'};

        /// <summary>
        ///     Is Visual Studio in design mode.
        /// </summary>
        /// <param name="serviceProvider">The service provider.</param>
        /// <returns>true if visual studio is in design mode</returns>
        public static bool IsVisualStudioInDesignMode(IServiceProvider site)
        {
            if (site == null)
            {
                throw new ArgumentNullException("site");
            }

            var selectionMonitor = site.GetService(typeof (IVsMonitorSelection)) as IVsMonitorSelection;
            uint cookie = 0;
            var active = 0;
            var designContext = VSConstants.UICONTEXT_DesignMode;
            ErrorHandler.ThrowOnFailure(selectionMonitor.GetCmdUIContextCookie(ref designContext, out cookie));
            ErrorHandler.ThrowOnFailure(selectionMonitor.IsCmdUIContextActive(cookie, out active));
            return active != 0;
        }

        /// <include file='doc\VsShellUtilities.uex' path='docs/doc[@for="Utilities.IsInAutomationFunction"]/*' />
        /// <devdoc>
        ///     Is an extensibility object executing an automation function.
        /// </devdoc>
        /// <param name="serviceProvider">The service provider.</param>
        /// <returns>true if the extensiblity object is executing an automation function.</returns>
        public static bool IsInAutomationFunction(IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                throw new ArgumentNullException("serviceProvider");
            }

            var extensibility = serviceProvider.GetService(typeof (IVsExtensibility)) as IVsExtensibility3;

            if (extensibility == null)
            {
                throw new InvalidOperationException();
            }
            var inAutomation = 0;
            ErrorHandler.ThrowOnFailure(extensibility.IsInAutomationFunction(out inAutomation));
            return inAutomation != 0;
        }

        /// <summary>
        ///     Creates a semicolon delinited list of strings. This can be used to provide the properties for
        ///     VSHPROPID_CfgPropertyPagesCLSIDList, VSHPROPID_PropertyPagesCLSIDList, VSHPROPID_PriorityPropertyPagesCLSIDList
        /// </summary>
        /// <param name="guids">An array of Guids.</param>
        /// <returns>A semicolon delimited string, or null</returns>
        [CLSCompliant(false)]
        public static string CreateSemicolonDelimitedListOfStringFromGuids(Guid[] guids)
        {
            if (guids == null || guids.Length == 0)
            {
                return null;
            }

            // Create a StringBuilder with a pre-allocated buffer big enough for the
            // final string. 39 is the length of a GUID in the "B" form plus the final ';'
            var stringList = new StringBuilder(39*guids.Length);
            for (var i = 0; i < guids.Length; i++)
            {
                stringList.Append(guids[i].ToString("B"));
                stringList.Append(";");
            }

            return stringList.ToString().TrimEnd(';');
        }

        /// <summary>
        ///     Take list of guids as a single string and generate an array of Guids from it
        /// </summary>
        /// <param name="guidList">Semi-colon separated list of Guids</param>
        /// <returns>Array of Guids</returns>
        [CLSCompliant(false)]
        public static Guid[] GuidsArrayFromSemicolonDelimitedStringOfGuids(string guidList)
        {
            if (guidList == null)
            {
                return null;
            }

            var guids = new List<Guid>();
            var guidsStrings = guidList.Split(';');
            foreach (var guid in guidsStrings)
            {
                if (!string.IsNullOrEmpty(guid))
                    guids.Add(new Guid(guid.Trim(curlyBraces)));
            }

            return guids.ToArray();
        }

        /// <summary>
        ///     Validates a file path by validating all file parts. If the
        ///     the file name is invalid it throws an exception if the project is in automation. Otherwise it shows a dialog box
        ///     with the error message.
        /// </summary>
        /// <param name="serviceProvider">The service provider</param>
        /// <param name="filePath">A full path to a file name</param>
        /// <exception cref="InvalidOperationException">In case of failure an InvalidOperationException is thrown.</exception>
        public static void ValidateFileName(IServiceProvider serviceProvider, string filePath)
        {
            var errorMessage = string.Empty;
            if (string.IsNullOrEmpty(filePath))
            {
                errorMessage = SR.GetString(SR.ErrorInvalidFileName, CultureInfo.CurrentUICulture);
            }
            else if (filePath.Length > NativeMethods.MAX_PATH)
            {
                errorMessage = string.Format(CultureInfo.CurrentCulture,
                    SR.GetString(SR.PathTooLong, CultureInfo.CurrentUICulture), filePath);
            }
            else if (ContainsInvalidFileNameChars(filePath))
            {
                errorMessage = SR.GetString(SR.ErrorInvalidFileName, CultureInfo.CurrentUICulture);
            }

            if (errorMessage.Length == 0)
            {
                var fileName = Path.GetFileName(filePath);
                if (string.IsNullOrEmpty(fileName) || IsFileNameInvalid(fileName))
                {
                    errorMessage = SR.GetString(SR.ErrorInvalidFileName, CultureInfo.CurrentUICulture);
                }
                else
                {
                    var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);

                    // If there is no filename or it starts with a leading dot issue an error message and quit.
                    if (string.IsNullOrEmpty(fileNameWithoutExtension) || fileNameWithoutExtension[0] == '.')
                    {
                        errorMessage = SR.GetString(SR.FileNameCannotContainALeadingPeriod, CultureInfo.CurrentUICulture);
                    }
                }
            }

            if (errorMessage.Length > 0)
            {
                // If it is not called from an automation method show a dialog box.
                if (!IsInAutomationFunction(serviceProvider))
                {
                    string title = null;
                    var icon = OLEMSGICON.OLEMSGICON_CRITICAL;
                    var buttons = OLEMSGBUTTON.OLEMSGBUTTON_OK;
                    var defaultButton = OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST;
                    VsShellUtilities.ShowMessageBox(serviceProvider, title, errorMessage, icon, buttons, defaultButton);
                }
                else
                {
                    throw new InvalidOperationException(errorMessage);
                }
            }
        }

        /// <summary>
        ///     Creates a CALPOLESTR from a list of strings
        ///     It is the responsability of the caller to release this memory.
        /// </summary>
        /// <param name="guids"></param>
        /// <returns>A CALPOLESTR that was created from the the list of strings.</returns>
        [CLSCompliant(false)]
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "CALPOLESTR")]
        public static CALPOLESTR CreateCALPOLESTR(IList<string> strings)
        {
            var calpolStr = new CALPOLESTR();

            if (strings != null)
            {
                // Demand unmanaged permissions in order to access unmanaged memory.
                new SecurityPermission(SecurityPermissionFlag.UnmanagedCode).Demand();

                calpolStr.cElems = (uint) strings.Count;

                var size = Marshal.SizeOf(typeof (IntPtr));

                calpolStr.pElems = Marshal.AllocCoTaskMem(strings.Count*size);

                var ptr = calpolStr.pElems;

                foreach (var aString in strings)
                {
                    var tempPtr = Marshal.StringToCoTaskMemUni(aString);
                    Marshal.WriteIntPtr(ptr, tempPtr);
                    ptr = new IntPtr(ptr.ToInt64() + size);
                }
            }

            return calpolStr;
        }

        /// <summary>
        ///     Creates a CADWORD from a list of tagVsSccFilesFlags. Memory is allocated for the elems.
        ///     It is the responsability of the caller to release this memory.
        /// </summary>
        /// <param name="guids"></param>
        /// <returns>A CADWORD created from the list of tagVsSccFilesFlags.</returns>
        [CLSCompliant(false)]
        [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "CADWORD")]
        public static CADWORD CreateCADWORD(IList<tagVsSccFilesFlags> flags)
        {
            var cadWord = new CADWORD();

            if (flags != null)
            {
                // Demand unmanaged permissions in order to access unmanaged memory.
                new SecurityPermission(SecurityPermissionFlag.UnmanagedCode).Demand();

                cadWord.cElems = (uint) flags.Count;

                var size = Marshal.SizeOf(typeof (uint));

                cadWord.pElems = Marshal.AllocCoTaskMem(flags.Count*size);

                var ptr = cadWord.pElems;

                foreach (var flag in flags)
                {
                    Marshal.WriteInt32(ptr, (int) flag);
                    ptr = new IntPtr(ptr.ToInt64() + size);
                }
            }

            return cadWord;
        }

        /// <summary>
        ///     Splits a bitmap from a Stream into an ImageList
        /// </summary>
        /// <param name="imageStream">A Stream representing a Bitmap</param>
        /// <returns>An ImageList object representing the images from the given stream</returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        public static ImageList GetImageList(Stream imageStream)
        {
            var ilist = new ImageList();

            if (imageStream == null)
            {
                return ilist;
            }
            ilist.ColorDepth = ColorDepth.Depth24Bit;
            ilist.ImageSize = new Size(16, 16);
            var bitmap = new Bitmap(imageStream);
            ilist.Images.AddStrip(bitmap);
            ilist.TransparentColor = Color.Magenta;
            return ilist;
        }

        /// <summary>
        ///     Splits a bitmap from a pointer to an ImageList
        /// </summary>
        /// <param name="imageListAsPointer">A pointer to a bitmap of images to split</param>
        /// <returns>An ImageList object representing the images from the given stream</returns>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        public static ImageList GetImageList(object imageListAsPointer)
        {
            ImageList images = null;

            var intPtr = new IntPtr((int) imageListAsPointer);
            var hImageList = new HandleRef(null, intPtr);
            var count = UnsafeNativeMethods.ImageList_GetImageCount(hImageList);

            if (count > 0)
            {
                // Create a bitmap big enough to hold all the images
                var b = new Bitmap(16*count, 16);
                var g = Graphics.FromImage(b);

                // Loop through and extract each image from the imagelist into our own bitmap
                var hDC = IntPtr.Zero;
                try
                {
                    hDC = g.GetHdc();
                    var handleRefDC = new HandleRef(null, hDC);
                    for (var i = 0; i < count; i++)
                    {
                        UnsafeNativeMethods.ImageList_Draw(hImageList, i, handleRefDC, i*16, 0, NativeMethods.ILD_NORMAL);
                    }
                }
                finally
                {
                    if (g != null && hDC != IntPtr.Zero)
                    {
                        g.ReleaseHdc(hDC);
                    }
                }

                // Create a new imagelist based on our stolen images
                images = new ImageList();
                images.ColorDepth = ColorDepth.Depth24Bit;
                images.ImageSize = new Size(16, 16);
                images.Images.AddStrip(b);
            }
            return images;
        }

        /// <summary>
        ///     Gets the active configuration name.
        /// </summary>
        /// <param name="automationObject">The automation object.</param>
        /// <returns>The name of the active configuartion.</returns>
        internal static string GetActiveConfigurationName(EnvDTE.Project automationObject)
        {
            if (automationObject == null)
            {
                throw new ArgumentNullException("automationObject");
            }

            var currentConfigName = string.Empty;
            if (automationObject.ConfigurationManager != null)
            {
                var activeConfig = automationObject.ConfigurationManager.ActiveConfiguration;
                if (activeConfig != null)
                {
                    currentConfigName = activeConfig.ConfigurationName;
                }
            }
            return currentConfigName;
        }

        /// <summary>
        ///     Verifies that two objects represent the same instance of a COM object.
        ///     This essentially compares the IUnkown pointers of the 2 objects.
        ///     This is needed in scenario where aggregation is involved.
        /// </summary>
        /// <param name="obj1">Can be an object, interface or IntPtr</param>
        /// <param name="obj2">Can be an object, interface or IntPtr</param>
        /// <returns>True if the 2 items represent the same thing</returns>
        [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "obj")]
        public static bool IsSameComObject(object obj1, object obj2)
        {
            var isSame = false;
            var unknown1 = IntPtr.Zero;
            var unknown2 = IntPtr.Zero;
            try
            {
                // If we have 2 null, then they are not COM objects and as such "it's not the same COM object"
                if (obj1 != null && obj2 != null)
                {
                    unknown1 = QueryInterfaceIUnknown(obj1);
                    unknown2 = QueryInterfaceIUnknown(obj2);

                    isSame = Equals(unknown1, unknown2);
                }
            }
            finally
            {
                if (unknown1 != IntPtr.Zero)
                {
                    Marshal.Release(unknown1);
                }

                if (unknown2 != IntPtr.Zero)
                {
                    Marshal.Release(unknown2);
                }
            }

            return isSame;
        }

        /// <summary>
        ///     Retrieve the IUnknown for the managed or COM object passed in.
        /// </summary>
        /// <param name="objToQuery">Managed or COM object.</param>
        /// <returns>Pointer to the IUnknown interface of the object.</returns>
        internal static IntPtr QueryInterfaceIUnknown(object objToQuery)
        {
            var releaseIt = false;
            var unknown = IntPtr.Zero;
            IntPtr result;
            try
            {
                if (objToQuery is IntPtr)
                {
                    unknown = (IntPtr) objToQuery;
                }
                else
                {
                    // This is a managed object (or RCW)
                    unknown = Marshal.GetIUnknownForObject(objToQuery);
                    releaseIt = true;
                }

                // We might already have an IUnknown, but if this is an aggregated
                // object, it may not be THE IUnknown until we QI for it.				
                var IID_IUnknown = VSConstants.IID_IUnknown;
                ErrorHandler.ThrowOnFailure(Marshal.QueryInterface(unknown, ref IID_IUnknown, out result));
            }
            finally
            {
                if (releaseIt && unknown != IntPtr.Zero)
                {
                    Marshal.Release(unknown);
                }
            }

            return result;
        }

        /// <summary>
        ///     Returns true if thename that can represent a path, absolut or relative, or a file name contains invalid filename
        ///     characters.
        /// </summary>
        /// <param name="name">File name</param>
        /// <returns>true if file name is invalid</returns>
        [SuppressMessage("Microsoft.Design", "CA1062:Validate arguments of public methods", MessageId = "0",
            Justification = "The name is validated.")]
        public static bool ContainsInvalidFileNameChars(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return true;
            }

            try
            {
                if (Path.IsPathRooted(name) && !name.StartsWith(@"\\", StringComparison.Ordinal))
                {
                    var root = Path.GetPathRoot(name);
                    name = name.Substring(root.Length);
                }
            }
                // The Path methods used by ContainsInvalidFileNameChars return argument exception if the filePath contains invalid characters.
            catch (ArgumentException)
            {
                return true;
            }

            var uri = new Url(name);

            // This might be confusing bur Url.IsFile means that the uri represented by the name is either absolut or relative.
            if (uri.IsFile)
            {
                var segments = uri.Segments;
                if (segments != null && segments.Length > 0)
                {
                    foreach (var segment in segments)
                    {
                        if (IsFilePartInValid(segment))
                        {
                            return true;
                        }
                    }

                    // Now the last segment should be specially taken care, since that cannot be all dots or spaces.
                    var lastSegment = segments[segments.Length - 1];
                    var filePart = Path.GetFileNameWithoutExtension(lastSegment);
                    if (IsFileNameAllGivenCharacter('.', filePart) || IsFileNameAllGivenCharacter(' ', filePart))
                    {
                        return true;
                    }
                }
            }
            else
            {
                // The assumption here is that we got a file name.
                var filePart = Path.GetFileNameWithoutExtension(name);
                if (IsFileNameAllGivenCharacter('.', filePart) || IsFileNameAllGivenCharacter(' ', filePart))
                {
                    return true;
                }


                return IsFilePartInValid(name);
            }

            return false;
        }

        /// Cehcks if a file name is valid.
        /// </devdoc>
        /// <param name="fileName">The name of the file</param>
        /// <returns>True if the file is valid.</returns>
        public static bool IsFileNameInvalid(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                return true;
            }

            if (IsFileNameAllGivenCharacter('.', fileName) || IsFileNameAllGivenCharacter(' ', fileName))
            {
                return true;
            }


            return IsFilePartInValid(fileName);
        }

        /// <summary>
        ///     Initializes the in memory project. Sets BuildEnabled on the project to true.
        /// </summary>
        /// <param name="engine">The build engine to use to create a build project.</param>
        /// <param name="fullProjectPath">The full path of the project.</param>
        /// <returns>A loaded msbuild project.</returns>
        internal static MSBuild.Project InitializeMsBuildProject(MSBuild.ProjectCollection buildEngine,
            string fullProjectPath)
        {
            if (string.IsNullOrEmpty(fullProjectPath))
            {
                throw new ArgumentException(SR.GetString(SR.InvalidParameter, CultureInfo.CurrentUICulture),
                    "fullProjectPath");
            }

            // Call GetFullPath to expand any relative path passed into this method.
            fullProjectPath = Path.GetFullPath(fullProjectPath);


            // Check if the project already has been loaded with the fullProjectPath. If yes return the build project associated to it.
            var loadedProject = new List<MSBuild.Project>(buildEngine.GetLoadedProjects(fullProjectPath));
            var buildProject = loadedProject != null && loadedProject.Count > 0 && loadedProject[0] != null
                ? loadedProject[0]
                : null;

            if (buildProject == null)
            {
                buildProject = buildEngine.LoadProject(fullProjectPath);
            }

            return buildProject;
        }

        /// <summary>
        ///     Loads a project file for the file. If the build project exists and it was loaded with a different file then it is
        ///     unloaded first.
        /// </summary>
        /// <param name="engine">The build engine to use to create a build project.</param>
        /// <param name="fullProjectPath">The full path of the project.</param>
        /// <param name="exitingBuildProject">An Existing build project that will be reloaded.</param>
        /// <returns>A loaded msbuild project.</returns>
        internal static MSBuild.Project ReinitializeMsBuildProject(MSBuild.ProjectCollection buildEngine,
            string fullProjectPath, MSBuild.Project exitingBuildProject)
        {
            // If we have a build project that has been loaded with another file unload it.
            try
            {
                if (exitingBuildProject != null && exitingBuildProject.ProjectCollection != null &&
                    !NativeMethods.IsSamePath(exitingBuildProject.FullPath, fullProjectPath))
                {
                    buildEngine.UnloadProject(exitingBuildProject);
                }
            }
                // We  catch Invalid operation exception because if the project was unloaded while we touch the ParentEngine the msbuild API throws. 
                // Is there a way to figure out that a project was unloaded?
            catch (InvalidOperationException)
            {
            }

            return InitializeMsBuildProject(buildEngine, fullProjectPath);
        }

        /// <summary>
        ///     Initialize the build engine. Sets the build enabled property to true. The engine is initialzed if the passed in
        ///     engine is null or does not have its bin path set.
        /// </summary>
        /// <param name="engine">An instance of MSBuild.ProjectCollection build engine, that will be checked if initialized.</param>
        /// <param name="engine">The service provider.</param>
        /// <returns>The buildengine to use.</returns>
        internal static MSBuild.ProjectCollection InitializeMsBuildEngine(MSBuild.ProjectCollection existingEngine,
            IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                throw new ArgumentNullException("serviceProvider");
            }

            if (existingEngine == null)
            {
                var buildEngine = MSBuild.ProjectCollection.GlobalProjectCollection;
                return buildEngine;
            }

            return existingEngine;
        }

        /// <summary>
        ///     Get the outer T implementation
        /// </summary>
        internal static T GetOuterAs<T>(object o)
            where T : class
        {
            T hierarchy = null;

            // The hierarchy of a node is its project node hierarchy.
            var projectUnknown = Marshal.GetIUnknownForObject(o);

            try
            {
                hierarchy = (T) Marshal.GetTypedObjectForIUnknown(projectUnknown, typeof (T));
            }
            finally
            {
                if (projectUnknown != IntPtr.Zero)
                {
                    Marshal.Release(projectUnknown);
                }
            }

            return hierarchy;
        }

        /// <summary>
        ///     >
        ///     Checks if the file name is all the given character.
        /// </summary>
        private static bool IsFileNameAllGivenCharacter(char c, string fileName)
        {
            // A valid file name cannot be all "c" .
            var charFound = 0;
            for (charFound = 0; charFound < fileName.Length && fileName[charFound] == c; ++charFound) ;
            if (charFound >= fileName.Length)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        ///     Checks whether a file part contains valid characters. The file part can be any part of a non rooted path.
        /// </summary>
        /// <param name="filePart"></param>
        /// <returns></returns>
        private static bool IsFilePartInValid(string filePart)
        {
            if (string.IsNullOrEmpty(filePart))
            {
                return true;
            }
            var reservedName = "(\\b(nul|con|aux|prn)\\b)|(\\b((com|lpt)[0-9])\\b)";
            var invalidChars = @"([/?:&\\*<>|#%" + '\"' + "])";
            var regexToUseForFileName = reservedName + "|" + invalidChars;
            var fileNameToVerify = filePart;

            // Define a regular expression that covers all characters that are not in the safe character sets.
            // It is compiled for performance.

            // The filePart might still be a file and extension. If it is like that then we must check them separately, since different rules apply
            var extension = string.Empty;
            try
            {
                extension = Path.GetExtension(filePart);
            }
                // We catch the ArgumentException because we want this method to return true if the filename is not valid. FilePart could be for example #�&%"�&"% and that would throw ArgumentException on GetExtension
            catch (ArgumentException)
            {
                return true;
            }

            if (!string.IsNullOrEmpty(extension))
            {
                // Check the extension first
                var regexToUseForExtension = invalidChars;
                var unsafeCharactersRegex = new Regex(regexToUseForExtension,
                    RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                var isMatch = unsafeCharactersRegex.IsMatch(extension);
                if (isMatch)
                {
                    return isMatch;
                }

                // We want to verify here everything but the extension.
                // We cannot use GetFileNameWithoutExtension because it might be that for example (..\\filename.txt) is passed in asnd that should fail, since that is not a valid filename.
                fileNameToVerify = filePart.Substring(0, filePart.Length - extension.Length);

                if (string.IsNullOrEmpty(fileNameToVerify))
                {
                    return true;
                }
            }

            // We verify CLOCK$ outside the regex since for some reason the regex is not matching the clock\\$ added.
            if (string.Compare(fileNameToVerify, "CLOCK$", StringComparison.OrdinalIgnoreCase) == 0)
            {
                return true;
            }

            var unsafeFileNameCharactersRegex = new Regex(regexToUseForFileName,
                RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return unsafeFileNameCharactersRegex.IsMatch(fileNameToVerify);
        }

        /// <summary>
        ///     Copy a directory recursively to the specified non-existing directory
        /// </summary>
        /// <param name="source">Directory to copy from</param>
        /// <param name="target">Directory to copy to</param>
        public static void RecursivelyCopyDirectory(string source, string target)
        {
            // Make sure it doesn't already exist
            if (Directory.Exists(target))
                throw new ArgumentException(string.Format(CultureInfo.CurrentCulture,
                    SR.GetString(SR.FileOrFolderAlreadyExists, CultureInfo.CurrentUICulture), target));

            Directory.CreateDirectory(target);
            var directory = new DirectoryInfo(source);

            // Copy files
            foreach (var file in directory.GetFiles())
            {
                file.CopyTo(Path.Combine(target, file.Name));
            }

            // Now recurse to child directories
            foreach (var child in directory.GetDirectories())
            {
                RecursivelyCopyDirectory(child.FullName, Path.Combine(target, child.Name));
            }
        }

        /// <summary>
        ///     Canonicalizes a file name, including:
        ///     - determines the full path to the file
        ///     - casts to upper case
        ///     Canonicalizing a file name makes it possible to compare file names using simple simple string comparison.
        ///     Note: this method does not handle shared drives and UNC drives.
        /// </summary>
        /// <param name="anyFileName">A file name, which can be relative/absolute and contain lower-case/upper-case characters.</param>
        /// <returns>Canonicalized file name.</returns>
        internal static string CanonicalizeFileName(string anyFileName)
        {
            // Get absolute path
            // Note: this will not handle UNC paths
            var fileInfo = new FileInfo(anyFileName);
            var fullPath = fileInfo.FullName;

            // Cast to upper-case
            fullPath = fullPath.ToUpper(CultureInfo.CurrentCulture);

            return fullPath;
        }

        /// <summary>
        ///     Determines if a file is a template.
        /// </summary>
        /// <param name="fileName">The file to check whether it is a template file</param>
        /// <returns>true if the file is a template file</returns>
        internal static bool IsTemplateFile(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                return false;
            }

            var extension = Path.GetExtension(fileName);
            return string.Compare(extension, ".vstemplate", StringComparison.OrdinalIgnoreCase) == 0 ||
                   string.Compare(extension, ".vsz", StringComparison.OrdinalIgnoreCase) == 0;
        }
    }
}