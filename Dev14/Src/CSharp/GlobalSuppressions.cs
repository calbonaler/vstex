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

using System.Diagnostics.CodeAnalysis;

// This file is used by Code Analysis to maintain SuppressMessage
// attributes that are applied to this project. Project-level
// suppressions either have no target or are given a specific target
// and scoped to a namespace, type, member, etc.
//
// To add a suppression to this file, right-click the message in the
// Error List, point to "Suppress Message(s)", and click "In Project
// Suppression File". You do not need to add suppressions to this
// file manually.

// The below contains suppressions from LinkDemand occurring from calls to the Marshal type's methods.
// Calls to these methods are absolute necessary to do successful VS integration. 
// All these methods are called from the VS Kernel, through interface calls that are expanded in the below methods. 
// Some of the below methods are non public but their caller's then are methods that are VS Integration interface implementations.
// The methods in which the LinkDemands were suppressed should be revisited though if adding further LinkDemands or directly Demanding UnmanagedCode Permission is a better choice.

//Interface parameters naming matching suppressions

[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#AdviseHierarchyEvents(Microsoft.VisualStudio.Shell.Interop.IVsHierarchyEvents,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#AdviseHierarchyEvents(Microsoft.VisualStudio.Shell.Interop.IVsHierarchyEvents,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#AdviseBuildStatusCallback(Microsoft.VisualStudio.Shell.Interop.IVsBuildStatusCallback,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartBuild(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartBuild(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartClean(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartUpToDateCheck(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartClean(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartUpToDateCheck(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartUpToDateCheck(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStatus(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#StartBuild(Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#StartBuild(Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#StartClean(Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#UnadviseBuildStatusCallback(System.UInt32)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#Stop(System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#StartUpToDateCheck(Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#StartUpToDateCheck(Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#StartClean(Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.BuildDependency.#get_HelpContext(System.UInt32&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.BuildDependency.#get_Description(System.String&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.BuildDependency.#get_CanonicalName(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#Wait(System.UInt32,System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.BuildDependency.#get_MustUpdateBefore(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.BuildDependency.#get_HelpFile(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterLoadProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterOpeningChildren(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterLoadProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterClosingChildren(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterOpenProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterOpenProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterRenameProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnBeforeCloseProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnBeforeCloseProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnBeforeClosingChildren(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnBeforeOpeningChildren(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnBeforeUnloadProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnBeforeUnloadProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnQueryChangeProjectParent(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnQueryChangeProjectParent(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnQueryChangeProjectParent(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnQueryCloseProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnQueryCloseProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnQueryCloseProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.SolutionListener.#OnQueryCloseSolution(System.Object,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnQueryUnloadProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListenerForProjectOpen.#OnAfterOpenSolution(System.Object,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListenerForProjectOpen.#OnAfterOpenSolution(System.Object,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#DeleteItem(System.UInt32,System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#DeleteItem(System.UInt32,System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#Exec(System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#Exec(System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#Exec(System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommand(System.UInt32,System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommand(System.UInt32,System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommand(System.UInt32,System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommand(System.UInt32,System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "5#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommand(System.UInt32,System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommand(System.UInt32,System.Guid&,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetCanonicalName(System.UInt32,System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetCanonicalName(System.UInt32,System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetGuidProperty(System.UInt32,System.Int32,System.Guid&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetGuidProperty(System.UInt32,System.Int32,System.Guid&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetNestedHierarchy(System.UInt32,System.Guid&,System.IntPtr&,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetNestedHierarchy(System.UInt32,System.Guid&,System.IntPtr&,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetProperty(System.UInt32,System.Int32,System.Object&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetProperty(System.UInt32,System.Int32,System.Object&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetProperty(System.UInt32,System.Int32,System.Object&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetSite(Microsoft.VisualStudio.OLE.Interop.IServiceProvider&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#IgnoreItemFileChanges(System.UInt32,System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#IgnoreItemFileChanges(System.UInt32,System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#IsItemDirty(System.UInt32,System.IntPtr,System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#IsItemDirty(System.UInt32,System.IntPtr,System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#IsItemDirty(System.UInt32,System.IntPtr,System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#IsItemReloadable(System.UInt32,System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#IsItemReloadable(System.UInt32,System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#OnBeforeDropNotify(Microsoft.VisualStudio.OLE.Interop.IDataObject,System.UInt32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#ParseCanonicalName(System.String,System.UInt32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#ParseCanonicalName(System.String,System.UInt32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryClose(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryDeleteItem(System.UInt32,System.UInt32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryDeleteItem(System.UInt32,System.UInt32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryDeleteItem(System.UInt32,System.UInt32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryStatus(System.Guid&,System.UInt32,Microsoft.VisualStudio.OLE.Interop.OLECMD[],System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryStatusCommand(System.UInt32,System.Guid&,System.UInt32,Microsoft.VisualStudio.OLE.Interop.OLECMD[],System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryStatusCommand(System.UInt32,System.Guid&,System.UInt32,Microsoft.VisualStudio.OLE.Interop.OLECMD[],System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryStatusCommand(System.UInt32,System.Guid&,System.UInt32,Microsoft.VisualStudio.OLE.Interop.OLECMD[],System.IntPtr)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#ReloadItem(System.UInt32,System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#ReloadItem(System.UInt32,System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SetGuidProperty(System.UInt32,System.Int32,System.Guid&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SetGuidProperty(System.UInt32,System.Int32,System.Guid&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SetProperty(System.UInt32,System.Int32,System.Object)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SetProperty(System.UInt32,System.Int32,System.Object)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#SetSite(Microsoft.VisualStudio.OLE.Interop.IServiceProvider)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#UnadviseHierarchyEvents(System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems.#AddFolder(System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems.#AddFolder(System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems.#AddFromDirectory(System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems.#AddFromFile(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems.#AddFromFileCopy(System.String)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems.#AddFromTemplate(System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems.#AddFromTemplate(System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.Automation.OAProjectItem`1.#Save(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.Automation.OAProjectItem`1.#SaveAs(System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#EnumOutputs(Microsoft.VisualStudio.Shell.Interop.IVsEnumOutputs&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_BuildableProjectCfg(Microsoft.VisualStudio.Shell.Interop.IVsBuildableProjectCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_CanonicalName(System.String&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_DisplayName(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_IsDebugOnly(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_IsPackaged(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_IsReleaseOnly(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_IsSpecifyingOutputSupported(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_Platform(System.Guid&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_ProjectCfgProvider(Microsoft.VisualStudio.Shell.Interop.IVsProjectCfgProvider&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_RootURL(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_TargetCodePage(System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#get_UpdateSequenceNumber(Microsoft.VisualStudio.OLE.Interop.ULARGE_INTEGER[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#GetCfg(Microsoft.VisualStudio.Shell.Interop.IVsCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#GetPages(Microsoft.VisualStudio.OLE.Interop.CAUUID[])")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#GetProjectDesignerPages(Microsoft.VisualStudio.OLE.Interop.CAUUID[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#GetProjectItem(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy&,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#GetProjectItem(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy&,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#OpenOutput(System.String,Microsoft.VisualStudio.Shell.Interop.IVsOutput&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#OpenOutput(System.String,Microsoft.VisualStudio.Shell.Interop.IVsOutput&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#QueryDebugLaunch(System.UInt32,System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#QueryDebugLaunch(System.UInt32,System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterAsynchOpenProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterAsynchOpenProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterChangeProjectParent(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.SolutionListener.#OnAfterCloseSolution(System.Object)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#AdviseBuildStatusCallback(Microsoft.VisualStudio.Shell.Interop.IVsBuildStatusCallback,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.BuildDependency.#get_ReferredProject(System.Object&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartClean(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#QueryStartBuild(System.UInt32,System.Int32[],System.Int32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#get_ProjectCfg(Microsoft.VisualStudio.Shell.Interop.IVsProjectCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.BuildDependency.#get_Type(System.Guid&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1701:ResourceStringCompoundWordsShouldBeCasedCorrectly",
        MessageId = "dataset", Scope = "resource", Target = "VsTeXProject.VisualStudio.Project.resources")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1703:ResourceStringsShouldBeSpelledCorrectly", MessageId = "Intelli",
        Scope = "resource", Target = "VsTeXProject.VisualStudio.Project.SecurityWarningDialog.resources")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.FileNode.#AfterSaveItemAs(System.IntPtr,System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NestedProjectNode.#IsItemDirty(System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.NestedProjectNode.#OnChanged(System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.NestedProjectNode.#OnRequestEdit(System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NestedProjectNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NestedProjectNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NestedProjectNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NodeProperties.#GetCfgProvider(Microsoft.VisualStudio.Shell.Interop.IVsCfgProvider&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NodeProperties.#GetPages(Microsoft.VisualStudio.OLE.Interop.CAUUID[])")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NodeProperties.#GetProjectDesignerPages(Microsoft.VisualStudio.OLE.Interop.CAUUID[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NodeProperties.#GetProjectItem(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy&,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NodeProperties.#GetProjectItem(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy&,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OnBeforeDropNotify(Microsoft.VisualStudio.OLE.Interop.IDataObject,System.UInt32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#OnClear(System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#OnPaste(System.Int32,System.UInt32)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#OnPaste(System.Int32,System.UInt32)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#RegisterClipboardNotifications(System.Boolean)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SettingsPage.#GetPageInfo(Microsoft.VisualStudio.OLE.Interop.PROPPAGEINFO[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.SettingsPage.#Help(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.SettingsPage.#Move(Microsoft.VisualStudio.OLE.Interop.RECT[])")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.SettingsPage.#SetObjects(System.UInt32,System.Object[])")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.SettingsPage.#SetObjects(System.UInt32,System.Object[])")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SettingsPage.#SetPageSite(Microsoft.VisualStudio.OLE.Interop.IPropertyPageSite)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.SettingsPage.#Show(System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SettingsPage.#TranslateAccelerator(Microsoft.VisualStudio.OLE.Interop.MSG[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#CreateGeneratorInstance(System.String,System.Int32&,System.Int32&,System.Int32&,Microsoft.VisualStudio.Shell.Interop.IVsSingleFileGenerator&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#CreateGeneratorInstance(System.String,System.Int32&,System.Int32&,System.Int32&,Microsoft.VisualStudio.Shell.Interop.IVsSingleFileGenerator&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#CreateGeneratorInstance(System.String,System.Int32&,System.Int32&,System.Int32&,Microsoft.VisualStudio.Shell.Interop.IVsSingleFileGenerator&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#CreateGeneratorInstance(System.String,System.Int32&,System.Int32&,System.Int32&,Microsoft.VisualStudio.Shell.Interop.IVsSingleFileGenerator&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#CreateGeneratorInstance(System.String,System.Int32&,System.Int32&,System.Int32&,Microsoft.VisualStudio.Shell.Interop.IVsSingleFileGenerator&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#GetDefaultGenerator(System.String,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#GetDefaultGenerator(System.String,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#GetGeneratorInformation(System.String,System.Int32&,System.Int32&,System.Int32&,System.Guid&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#GetGeneratorInformation(System.String,System.Int32&,System.Int32&,System.Int32&,System.Guid&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#GetGeneratorInformation(System.String,System.Int32&,System.Int32&,System.Int32&,System.Guid&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#GetGeneratorInformation(System.String,System.Int32&,System.Int32&,System.Int32&,System.Guid&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SingleFileGeneratorFactory.#GetGeneratorInformation(System.String,System.Int32&,System.Int32&,System.Int32&,System.Guid&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NestedProjectNode.#IsItemDirty(System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SettingsPage.#Activate(System.IntPtr,Microsoft.VisualStudio.OLE.Interop.RECT[],System.Int32)"
        )]

//Ref arguments suppressions

[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "0#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.TokenProcessor.#ReplaceBetweenTokens(System.String&,VsTeXProject.VisualStudio.Project.ReplaceBetweenPairToken)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "0#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.TokenProcessor.#ReplaceTokens(System.String&,VsTeXProject.VisualStudio.Project.ReplacePairToken)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "3#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryStatusOnNode(System.Guid,System.UInt32,System.IntPtr,VsTeXProject.VisualStudio.Project.QueryStatusResult&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "3#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DocumentManager.#OpenWithSpecific(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "2#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.FileDocumentManager.#Open(System.Boolean,System.Boolean,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "1#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DocumentManager.#OpenWithSpecific(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "0#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.TokenProcessor.#DeleteTokens(System.String&,VsTeXProject.VisualStudio.Project.DeleteToken)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "0#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DocumentManager.#Open(System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "0#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DocumentManager.#CloseWindowFrame(Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "4#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Load(System.String,System.String,System.String,System.UInt32,System.Guid&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "1#", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#SetGuidProperty(System.Int32,System.Guid&)")]

//Out arguments suppressions

[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "2#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DocumentManager.#Open(System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "4#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.FileDocumentManager.#Open(System.Boolean,System.Boolean,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "3#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#DisableCommandOnNodesThatDoNotSupportMultiSelection(System.Guid,System.UInt32,System.Collections.Generic.IList`1<VisualStudio.Project.HierarchyNode>,System.Boolean&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "6#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommandIndependentOfSelection(System.Guid,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr,VsTeXProject.VisualStudio.Project.CommandOrigin,System.Boolean&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "7#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommandThatDependsOnSelectedNodes(System.Guid,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr,VsTeXProject.VisualStudio.Project.CommandOrigin,System.Collections.Generic.IList`1<VisualStudio.Project.HierarchyNode>,System.Boolean&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "2#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#QueryStatusCommandFromOleCommandTarget(System.Guid,System.UInt32,System.Boolean&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "1#", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetGuidProperty(System.Int32,System.Guid&)")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "5#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DocumentManager.#OpenWithSpecific(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "3#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.FileDocumentManager.#Open(System.Boolean,System.Boolean,System.Guid,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "0#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ReferenceNode.#CanAddReference(VsTeXProject.VisualStudio.Project.ReferenceNode+CannotAddReferenceErrorMessage&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "5#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Load(System.String,System.String,System.String,System.UInt32,System.Guid&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "7#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#ExecCommandThatDependsOnSelectedNodes(System.Guid,System.UInt32,System.UInt32,System.IntPtr,System.IntPtr,VsTeXProject.VisualStudio.Project.CommandOrigin,System.Collections.Generic.IList`1<VsTeXProject.VisualStudio.Project.HierarchyNode>,System.Boolean&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1021:AvoidOutParameters", MessageId = "3#", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#DisableCommandOnNodesThatDoNotSupportMultiSelection(System.Guid,System.UInt32,System.Collections.Generic.IList`1<VsTeXProject.VisualStudio.Project.HierarchyNode>,System.Boolean&)"
        )]

//Default constructors for COM suppressions

[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.AssemblyReferenceNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.BuildableProjectConfig")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ComReferenceNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ConnectionPointContainer")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.DependentFileNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.DependentFileNodeProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.FileNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.FileNodeProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.FolderNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.FolderNodeProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.NodeProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAAssemblyReference")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAComReference")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAFileItem")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAFolderItem")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OANestedProjectItem")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OANullProperty")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProject")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProjectItems")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProjectReference")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProperty")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceFolderItem")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceItem")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferences")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAVSProject")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAVSProjectEvents")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.OutputGroup")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectConfig")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectConfigProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectContainerNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNodeProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectPackage")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectReferenceNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectReferencesProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ReferenceContainerNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ReferenceNode")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ReferenceNodeProperties")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.SettingsPage")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.SingleFileGeneratorNodeProperties")]

//PInvoke suppressions

[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1414:MarkBooleanPInvokeArgumentsWithMarshalAs", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NativeMethods.#DestroyIcon(System.IntPtr)")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1414:MarkBooleanPInvokeArgumentsWithMarshalAs", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NativeMethods.#GetBinaryType(System.String,System.UInt32&)")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1414:MarkBooleanPInvokeArgumentsWithMarshalAs", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.UnsafeNativeMethods.#GlobalUnlock(System.Runtime.InteropServices.HandleRef)"
        )]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1414:MarkBooleanPInvokeArgumentsWithMarshalAs", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.UnsafeNativeMethods.#GlobalUnLock(System.IntPtr)")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1414:MarkBooleanPInvokeArgumentsWithMarshalAs", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.UnsafeNativeMethods.#ImageList_Draw(System.Runtime.InteropServices.HandleRef,System.Int32,System.Runtime.InteropServices.HandleRef,System.Int32,System.Int32,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1414:MarkBooleanPInvokeArgumentsWithMarshalAs", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.UnsafeNativeMethods.#SHGetPathFromIDList(System.IntPtr,System.IntPtr)")]
[assembly:
    SuppressMessage("Microsoft.Interoperability", "CA1414:MarkBooleanPInvokeArgumentsWithMarshalAs", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NativeMethods.#IsDialogMessageA(System.IntPtr,Microsoft.VisualStudio.OLE.Interop.MSG&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.UnsafeNativeMethods.#RegisterClipboardFormat(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.UnsafeNativeMethods.#GlobalSize(System.IntPtr)")]
[assembly:
    SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "return",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.UnsafeNativeMethods.#GlobalSize(System.Runtime.InteropServices.HandleRef)"
        )]

//'Flags' naming suppressions

[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ModuleKindFlags")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectNode.#VirtualProjectFlags")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetQueryRemoveFileFlags(System.String[])")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetQueryAddFileFlags(System.String[])")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetAddFileFlags(System.String[])")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Utilities.#CreateCADWORD(System.Collections.Generic.IList`1<Microsoft.VisualStudio.Shell.Interop.tagVsSccFilesFlags>)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Load(System.String,System.String,System.String,System.UInt32,System.Guid&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#AddNestedProjectFromTemplate(System.String,System.String,System.String,VsTeXProject.VisualStudio.Project.ProjectElement,Microsoft.VisualStudio.Shell.Interop.__VSCREATEPROJFLAGS)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#AddNestedProjectFromTemplate(VsTeXProject.VisualStudio.Project.ProjectElement,Microsoft.VisualStudio.Shell.Interop.__VSCREATEPROJFLAGS)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#AddExistingNestedProject(VsTeXProject.VisualStudio.Project.ProjectElement,Microsoft.VisualStudio.Shell.Interop.__VSCREATEPROJFLAGS)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NestedProjectNode.#Init(System.String,System.String,System.String,Microsoft.VisualStudio.Shell.Interop.__VSCREATEPROJFLAGS)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#OnPropertyChanged(VsTeXProject.VisualStudio.Project.HierarchyNode,System.Int32,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flag", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#IgnoreItemFileChanges(System.Boolean)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetSccSpecialFiles(System.String,System.Collections.Generic.IList`1<System.String>,System.Collections.Generic.IList`1<Microsoft.VisualStudio.Shell.Interop.tagVsSccFilesFlags>)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetSccFiles(System.Collections.Generic.IList`1<System.String>,System.Collections.Generic.IList`1<Microsoft.VisualStudio.Shell.Interop.tagVsSccFilesFlags>)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DocumentManager.#OpenWithSpecific(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flag", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DocumentManager.#Close(Microsoft.VisualStudio.Shell.Interop.__FRAMECLOSE)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1726:UsePreferredTerms", MessageId = "Flags", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetRemoveFileFlags(System.String[])")]

//Properties instead of methods suppressions

[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.DocumentManager.#GetFullPathForDocument()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.DocumentManager.#GetOwnerCaption()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetAutomationObject()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetDocumentManager()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetDragTargetHandlerNode()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetEditLabel()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetMkDocument()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetRelationalName()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetRelationNameExtension()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.IReferenceContainerProvider.#GetReferenceContainer()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems.#GetListOfProjectItems()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectElement.#GetFullPathForElement()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetCompiler()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetInner()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetOutputGroupNames()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectOptions.#GetOptionHelp()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetSelectedNodes()")]

//Generic nesting suppressions

[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.BuildableProjectConfig.#Build(System.UInt32,Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetOutputGroupNames()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#NewConfigProperties")]

//Methods that swallows exceptions suppressions

[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Web.CodeBehindCodeGenerator.#GetDesignerItem(VsTeXProject.VisualStudio.Project.Web.VsHierarchyItem,System.Boolean)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Web.CodeBehindCodeGenerator.#FindClass(VsTeXProject.VisualStudio.Project.Web.VsHierarchyItem,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Web.CodeBehindCodeGenerator.#DisposeGenerateState()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Web.CodeBehindCodeGenerator.#GetFieldNames(EnvDTE.CodeClass,System.Boolean,System.Boolean,System.Int32,System.Int32,VsTeXProject.VisualStudio.Project.Web.FieldDataDictionary&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Web.CodeBehindCodeGenerator.#GetFieldNames(System.String[],System.Boolean)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#Close()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NativeMethods+ConnectionPointCookie.#.ctor(System.Object,System.Object,System.Type,System.Boolean)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NativeMethods+DataStreamFromComStream.#Flush()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Web.VsHierarchyItem.#FullPath()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Web.VsHierarchyItem.#GetGuidPropHelper(Microsoft.VisualStudio.Shell.Interop.__VSHPROPID)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Web.VsHierarchyItem.#GetPropHelper(System.UInt32,System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Web.VsHierarchyItem.#GetService`1()")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Web.WAUtilities.#CreateInstance`1(System.IServiceProvider,System.Guid)")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Web.WAUtilities.#GetService`1(System.IServiceProvider)")]

//Local variables names matching instance variable names suppressions

[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "dropDataType", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#PasteFromClipboard(VsTeXProject.VisualStudio.Project.HierarchyNode)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "dropDataType", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Drop(Microsoft.VisualStudio.OLE.Interop.IDataObject,System.UInt32,System.UInt32,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "dropDataType", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#ProcessSelectionDataObject(Microsoft.VisualStudio.OLE.Interop.IDataObject,VsTeXProject.VisualStudio.Project.HierarchyNode)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "filename",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#InitializeForOuter(System.String,System.String,System.String,System.UInt32,System.Guid&,System.IntPtr&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "filename",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#SaveCompleted(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "isDirty",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#IsDirty(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "isDirty",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#CleanupSelectionDataObject(System.Boolean,System.Boolean,System.Boolean,System.Boolean)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "isDirty",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#IsFlavorDirty()")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "isDirty",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OnBeforeDropNotify(Microsoft.VisualStudio.OLE.Interop.IDataObject,System.UInt32,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "name",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.SettingsPage.#GetTypedConfigProperty(System.String,System.Type)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "name",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.SettingsPage.#GetTypedProperty(System.String,System.Type)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "name",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.SettingsPage.#SetConfigProperty(System.String,System.String)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "options",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetTargetPlatform(VsTeXProject.VisualStudio.Project.ProjectOptions)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "options",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetOutputAssembly(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "options",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetProjectOptions(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "options",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#PrepareBuild(System.String,System.Boolean)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "options",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#SetBuildConfigurationProperties(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "project",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#CreateOutputGroup(VsTeXProject.VisualStudio.Project.ProjectNode,System.Collections.Generic.KeyValuePair`2<System.String,System.String>)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "projectName", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NestedProjectNode.#Init(System.String,System.String,System.String,Microsoft.VisualStudio.Shell.Interop.__VSCREATEPROJFLAGS)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "sccAuxPath", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccLocation(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "sccAuxPath", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccSettings(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "sccLocalPath", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccLocation(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "sccLocalPath", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccSettings(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "sccProjectName", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccLocation(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "sccProjectName", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccSettings(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "sccProvider", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccLocation(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "sccProvider", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccSettings(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "site",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSite(Microsoft.VisualStudio.OLE.Interop.IServiceProvider)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames",
        MessageId = "dropDataType", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#QueryDropEffect(VsTeXProject.VisualStudio.Project.DropDataType,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "i",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.EnumSTATDATA.#Microsoft.VisualStudio.OLE.Interop.IEnumSTATDATA.Skip(System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Maintainability", "CA1500:VariableNamesShouldNotMatchFieldNames", MessageId = "i",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.EnumSTATDATA.#Microsoft.VisualStudio.OLE.Interop.IEnumSTATDATA.Next(System.UInt32,Microsoft.VisualStudio.OLE.Interop.STATDATA[],System.UInt32&)"
        )]

//COM Exceptions suppressions

[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceBase`1.#ExtenderCATID")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceBase`1.#Culture")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceBase`1.#ExtenderNames")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceBase`1.#Identity")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceBase`1.#Name")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceBase`1.#PublicKeyToken")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceBase`1.#Type")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OASolutionFolder`1.#Parent")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProjectItem`1.#Collection")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProjectItem`1.#IsDirty")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceBase`1.#CopyLocal")]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.DataObject.#Microsoft.VisualStudio.OLE.Interop.IDataObject.GetCanonicalFormatEtc(Microsoft.VisualStudio.OLE.Interop.FORMATETC[],Microsoft.VisualStudio.OLE.Interop.FORMATETC[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#.ctor(VsTeXProject.VisualStudio.Project.ProjectNode,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#RunVsTemplateWizard(VsTeXProject.VisualStudio.Project.ProjectElement,System.Boolean)"
        )]

//IEnumerable implementation suppressions

[assembly:
    SuppressMessage("Microsoft.Design", "CA1010:CollectionsShouldImplementGenericInterface", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferences")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1010:CollectionsShouldImplementGenericInterface", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProperties")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1010:CollectionsShouldImplementGenericInterface", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OANavigableProjectItems")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1010:CollectionsShouldImplementGenericInterface", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAProjectItems")]

//Other suppressions

[assembly:
    SuppressMessage("Microsoft.Design", "CA1034:NestedTypesShouldNotBeVisible", Scope = "type",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode+ImageName")]
[assembly:
    SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "trailer", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetProperty(System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "vspropId", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.HierarchyNode.#GetGuidProperty(System.UInt32,System.Int32,System.Guid&)")
]
[assembly:
    SuppressMessage("Microsoft.Performance", "CA1805:DoNotInitializeUnnecessarily", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NativeMethods+ConnectionPointCookie.#.ctor(System.Object,System.Object,System.Type,System.Boolean)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConnectionPointContainer.#Microsoft.VisualStudio.OLE.Interop.IConnectionPointContainer.FindConnectionPoint(System.Guid&,Microsoft.VisualStudio.OLE.Interop.IConnectionPoint&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NodeProperties.#EnvDTE80.IInternalExtenderProvider.CanExtend(System.String,System.String,System.Object)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NodeProperties.#EnvDTE80.IInternalExtenderProvider.GetExtender(System.String,System.String,System.Object,EnvDTE.IExtenderSite,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.NodeProperties.#EnvDTE80.IInternalExtenderProvider.GetExtenderNames(System.String,System.Object)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectConfig.#Microsoft.VisualStudio.Shell.Interop.IVsProjectFlavorCfg.Close()"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#VisualStudio.Project.IProjectEventsProvider.ProjectEventsProvider"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Microsoft.VisualStudio.Shell.Interop.IVsBuildPropertyStorage.GetItemAttribute(System.UInt32,System.String,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Microsoft.VisualStudio.Shell.Interop.IVsBuildPropertyStorage.GetPropertyValue(System.String,System.String,System.UInt32,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Microsoft.VisualStudio.Shell.Interop.IVsBuildPropertyStorage.RemoveProperty(System.String,System.String,System.UInt32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Microsoft.VisualStudio.Shell.Interop.IVsBuildPropertyStorage.SetItemAttribute(System.UInt32,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Microsoft.VisualStudio.Shell.Interop.IVsBuildPropertyStorage.SetPropertyValue(System.String,System.String,System.UInt32,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#Microsoft.VisualStudio.Shell.Interop.IVsProjectFlavorCfgProvider.CreateProjectFlavorCfg(Microsoft.VisualStudio.Shell.Interop.IVsCfg,Microsoft.VisualStudio.Shell.Interop.IVsProjectFlavorCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConnectionPointContainer.#Microsoft.VisualStudio.OLE.Interop.IConnectionPointContainer.EnumConnectionPoints(Microsoft.VisualStudio.OLE.Interop.IEnumConnectionPoints&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#VsTeXProject.VisualStudio.Project.IProjectEventsProvider.ProjectEventsProvider"
        )]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetMsBuildProperty(System.String,System.Boolean)")]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddNodeIfTargetExistInStorage(VsTeXProject.VisualStudio.Project.HierarchyNode,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddFolderFromOtherProject(System.String,VsTeXProject.VisualStudio.Project.HierarchyNode)"
        )]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectElement.#GetMetadataAndThrow(System.String,System.Exception)"
        )]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#GetRegisteredProject(VsTeXProject.VisualStudio.Project.ProjectElement)"
        )]
[assembly:
    SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectConfig.#GetMsBuildProperty(System.String,System.Boolean)")]
[assembly:
    SuppressMessage("Microsoft.Design", "CA1065:DoNotRaiseExceptionsInUnexpectedLocations", Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OAVSProject.#Imports")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectBuildDependency.#get_CanonicalName(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectBuildDependency.#get_HelpContext(System.UInt32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.SolutionListenerForProjectReferenceUpdate.#OnBeforeUnloadProject(Microsoft.VisualStudio.Shell.Interop.IVsHierarchy,Microsoft.VisualStudio.Shell.Interop.IVsHierarchy)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#TransferItem(System.String,System.String,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#TransferItem(System.String,System.String,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#TransferItem(System.String,System.String,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccLocation(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccLocation(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccLocation(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetSccLocation(System.String,System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#SetInnerProject(System.Object)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetHostObject(System.String,System.String,System.Object)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetHostObject(System.String,System.String,System.Object)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SetHostObject(System.String,System.String,System.Object)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#SetAggregateProjectTypeGuids(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SccGlyphChanged(System.Int32,System.UInt32[],Microsoft.VisualStudio.Shell.Interop.VsStateIcon[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SccGlyphChanged(System.Int32,System.UInt32[],Microsoft.VisualStudio.Shell.Interop.VsStateIcon[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SccGlyphChanged(System.Int32,System.UInt32[],Microsoft.VisualStudio.Shell.Interop.VsStateIcon[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#SccGlyphChanged(System.Int32,System.UInt32[],Microsoft.VisualStudio.Shell.Interop.VsStateIcon[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#SaveCompleted(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#Save(System.String,System.Int32,System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#Save(System.String,System.Int32,System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#Save(System.String,System.Int32,System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#ReopenItem(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#ReopenItem(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#ReopenItem(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "5#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#ReopenItem(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#ReopenItem(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#ReopenItem(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#RemoveItem(System.UInt32,System.UInt32,System.Int32&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#RemoveItem(System.UInt32,System.UInt32,System.Int32&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#RemoveItem(System.UInt32,System.UInt32,System.Int32&)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItemWithSpecific(System.UInt32,System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItemWithSpecific(System.UInt32,System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItemWithSpecific(System.UInt32,System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "6#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItemWithSpecific(System.UInt32,System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItemWithSpecific(System.UInt32,System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItemWithSpecific(System.UInt32,System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "5#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItemWithSpecific(System.UInt32,System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItem(System.UInt32,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItem(System.UInt32,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenItem(System.UInt32,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#OpenDependency(System.String,Microsoft.VisualStudio.Shell.Interop.IVsDependency&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#Load(System.String,System.UInt32,System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#Load(System.String,System.UInt32,System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#Load(System.String,System.UInt32,System.Int32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#IsDocumentInProject(System.String,System.Int32&,Microsoft.VisualStudio.Shell.Interop.VSDOCUMENTPRIORITY[],System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#IsDocumentInProject(System.String,System.Int32&,Microsoft.VisualStudio.Shell.Interop.VSDOCUMENTPRIORITY[],System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#IsDocumentInProject(System.String,System.Int32&,Microsoft.VisualStudio.Shell.Interop.VSDOCUMENTPRIORITY[],System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#IsDocumentInProject(System.String,System.Int32&,Microsoft.VisualStudio.Shell.Interop.VSDOCUMENTPRIORITY[],System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#IsDirty(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#InitNew(System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "5#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#InitializeForOuter(System.String,System.String,System.String,System.UInt32,System.Guid&,System.IntPtr&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#InitializeForOuter(System.String,System.String,System.String,System.UInt32,System.Guid&,System.IntPtr&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#InitializeForOuter(System.String,System.String,System.String,System.UInt32,System.Guid&,System.IntPtr&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#InitializeForOuter(System.String,System.String,System.String,System.UInt32,System.Guid&,System.IntPtr&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#InitializeForOuter(System.String,System.String,System.String,System.UInt32,System.Guid&,System.IntPtr&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#InitializeForOuter(System.String,System.String,System.String,System.UInt32,System.Guid&,System.IntPtr&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "6#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#InitializeForOuter(System.String,System.String,System.String,System.UInt32,System.Guid&,System.IntPtr&,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetSccSpecialFiles(System.UInt32,System.String,Microsoft.VisualStudio.OLE.Interop.CALPOLESTR[],Microsoft.VisualStudio.OLE.Interop.CADWORD[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetSccSpecialFiles(System.UInt32,System.String,Microsoft.VisualStudio.OLE.Interop.CALPOLESTR[],Microsoft.VisualStudio.OLE.Interop.CADWORD[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetSccSpecialFiles(System.UInt32,System.String,Microsoft.VisualStudio.OLE.Interop.CALPOLESTR[],Microsoft.VisualStudio.OLE.Interop.CADWORD[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetSccFiles(System.UInt32,Microsoft.VisualStudio.OLE.Interop.CALPOLESTR[],Microsoft.VisualStudio.OLE.Interop.CADWORD[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetSccFiles(System.UInt32,Microsoft.VisualStudio.OLE.Interop.CALPOLESTR[],Microsoft.VisualStudio.OLE.Interop.CADWORD[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetMkDocument(System.UInt32,System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetMkDocument(System.UInt32,System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetItemContext(System.UInt32,Microsoft.VisualStudio.OLE.Interop.IServiceProvider&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetItemContext(System.UInt32,Microsoft.VisualStudio.OLE.Interop.IServiceProvider&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetFormatList(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetFile(System.Int32,System.UInt32,System.UInt32&,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetFile(System.Int32,System.UInt32,System.UInt32&,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetFile(System.Int32,System.UInt32,System.UInt32&,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetFile(System.Int32,System.UInt32,System.UInt32&,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetCurFile(System.String&,System.UInt32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetCurFile(System.String&,System.UInt32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetClassID(System.Guid&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GetCfgProvider(Microsoft.VisualStudio.Shell.Interop.IVsCfgProvider&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetBuildSystemKind(System.UInt32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetAggregateProjectTypeGuids(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GenerateUniqueItemName(System.UInt32,System.String,System.String,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GenerateUniqueItemName(System.UInt32,System.String,System.String,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GenerateUniqueItemName(System.UInt32,System.String,System.String,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#GenerateUniqueItemName(System.UInt32,System.String,System.String,System.String&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#EnumDependencies(Microsoft.VisualStudio.Shell.Interop.IVsEnumDependencies&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#BuildTarget(System.String,System.Boolean&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#BuildTarget(System.String,System.Boolean&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "10#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "8#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "9#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "7#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "6#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "5#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItemWithSpecific(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,System.UInt32,System.Guid&,System.String,System.Guid&,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "6#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItem(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItem(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItem(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItem(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItem(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItem(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "5#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddItem(System.UInt32,Microsoft.VisualStudio.Shell.Interop.VSADDITEMOPERATION,System.String,System.UInt32,System.String[],System.IntPtr,Microsoft.VisualStudio.Shell.Interop.VSADDRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectNode.#AddComponent(Microsoft.VisualStudio.Shell.Interop.VSADDCOMPOPERATION,System.UInt32,System.IntPtr[],System.IntPtr,Microsoft.VisualStudio.Shell.Interop.VSADDCOMPRESULT[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "4#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#SaveItem(Microsoft.VisualStudio.Shell.Interop.VSSAVEFLAGS,System.String,System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#IsItemDirty(System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ProjectContainerNode.#IsItemDirty(System.UInt32,System.IntPtr,System.Int32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Automation.OASolutionFolder`1.#AddFromTemplate(System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Automation.OASolutionFolder`1.#AddFromTemplate(System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.Automation.OASolutionFolder`1.#AddFromTemplate(System.String,System.String,System.String)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.Automation.OASolutionFolder`1.#AddFromFile(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferenceItem.#Open(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.Automation.OAProject.#SaveAs(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.Automation.OAProject.#Save(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.Automation.OAFileItem.#SaveAs(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.Automation.OAFileItem.#Open(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectBuildDependency.#get_Type(System.Guid&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectBuildDependency.#get_ReferredProject(System.Object&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectBuildDependency.#get_MustUpdateBefore(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectBuildDependency.#get_HelpFile(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.NestedProjectBuildDependency.#get_Description(System.String&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "5#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.FileDocumentManager.#OpenWithSpecific(System.UInt32,System.Guid&,System.String,System.Guid&,System.IntPtr,Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame&,VsTeXProject.VisualStudio.Project.WindowFrameShowAction)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member", Target = "VsTeXProject.VisualStudio.Project.EnumDependencies.#Skip(System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.EnumDependencies.#Next(System.UInt32,Microsoft.VisualStudio.Shell.Interop.IVsDependency[],System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.EnumDependencies.#Next(System.UInt32,Microsoft.VisualStudio.Shell.Interop.IVsDependency[],System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.EnumDependencies.#Next(System.UInt32,Microsoft.VisualStudio.Shell.Interop.IVsDependency[],System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.EnumDependencies.#Clone(Microsoft.VisualStudio.Shell.Interop.IVsEnumDependencies&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#UnadviseCfgProviderEvents(System.UInt32)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#RenameCfgsOfCfgName(System.String,System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#RenameCfgsOfCfgName(System.String,System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#OpenProjectCfg(System.String,Microsoft.VisualStudio.Shell.Interop.IVsProjectCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#OpenProjectCfg(System.String,Microsoft.VisualStudio.Shell.Interop.IVsProjectCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetSupportedPlatformNames(System.UInt32,System.String[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetSupportedPlatformNames(System.UInt32,System.String[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetPlatformNames(System.UInt32,System.String[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetPlatformNames(System.UInt32,System.String[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "3#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgs(System.UInt32,Microsoft.VisualStudio.Shell.Interop.IVsCfg[],System.UInt32[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgs(System.UInt32,Microsoft.VisualStudio.Shell.Interop.IVsCfg[],System.UInt32[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgs(System.UInt32,Microsoft.VisualStudio.Shell.Interop.IVsCfg[],System.UInt32[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgProviderProperty(System.Int32,System.Object&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgOfName(System.String,System.String,Microsoft.VisualStudio.Shell.Interop.IVsCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgOfName(System.String,System.String,Microsoft.VisualStudio.Shell.Interop.IVsCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgOfName(System.String,System.String,Microsoft.VisualStudio.Shell.Interop.IVsCfg&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgNames(System.UInt32,System.String[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "2#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetCfgNames(System.UInt32,System.String[],System.UInt32[])"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetAutomationObject(System.String,System.Object&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#GetAutomationObject(System.String,System.Object&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#get_UsesIndependentConfigurations(System.Int32&)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#DeleteCfgsOfPlatformName(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#DeleteCfgsOfCfgName(System.String)")]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#AdviseCfgProviderEvents(Microsoft.VisualStudio.Shell.Interop.IVsCfgProviderEvents,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#AdviseCfgProviderEvents(Microsoft.VisualStudio.Shell.Interop.IVsCfgProviderEvents,System.UInt32&)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#AddCfgsOfPlatformName(System.String,System.String)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target = "VsTeXProject.VisualStudio.Project.ConfigProvider.#AddCfgsOfPlatformName(System.String,System.String)")
]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "1#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#AddCfgsOfCfgName(System.String,System.String,System.Int32)"
        )]
[assembly:
    SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#",
        Scope = "member",
        Target =
            "VsTeXProject.VisualStudio.Project.ConfigProvider.#AddCfgsOfCfgName(System.String,System.String,System.Int32)"
        )]
[assembly: SuppressMessage("Microsoft.Naming", "CA1725:ParameterNamesShouldMatchBaseDeclaration", MessageId = "0#", Scope = "member", Target = "VsTeXProject.VisualStudio.Project.Automation.OAReferences.#AddProject(EnvDTE.Project)")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Te")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Te", Scope = "namespace", Target = "VsTeXProject.VisualStudio.Project")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Te", Scope = "namespace", Target = "VsTeXProject.VisualStudio.Project.Automation")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Te", Scope = "type", Target = "VsTeXProject.VisualStudio.Project.TeXProcessor")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Te", Scope = "type", Target = "VsTeXProject.VisualStudio.Project.TeXProcessorConverter")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Vs")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Vs", Scope = "namespace", Target = "VsTeXProject.VisualStudio.Project")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "Vs", Scope = "namespace", Target = "VsTeXProject.VisualStudio.Project.Automation")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type", Target = "VsTeXProject.VisualStudio.Project.ComReferenceProperties")]
[assembly: SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ReferenceContainerNode.#CreateReferenceNode(Microsoft.VisualStudio.Shell.Interop.VSCOMPONENTSELECTORDATA,System.String)")]
[assembly: SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ReferenceContainerNode.#CreateFileComponent(Microsoft.VisualStudio.Shell.Interop.VSCOMPONENTSELECTORDATA,System.String)")]
[assembly: SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ReferenceContainerNode.#CreateComReferenceNode(Microsoft.VisualStudio.Shell.Interop.VSCOMPONENTSELECTORDATA,System.String)")]
[assembly: SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ReferenceContainerNode.#AddReferenceFromSelectorData(Microsoft.VisualStudio.Shell.Interop.VSCOMPONENTSELECTORDATA,System.String)")]
[assembly: SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ProjectNode.#GetProjectOptions(System.String)")]
[assembly: SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Scope = "member", Target = "VsTeXProject.VisualStudio.Project.IReferenceContainer.#AddReferenceFromSelectorData(Microsoft.VisualStudio.Shell.Interop.VSCOMPONENTSELECTORDATA,System.String)")]
[assembly: SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Scope = "member", Target = "VsTeXProject.VisualStudio.Project.ComReferenceNode.#.ctor(VsTeXProject.VisualStudio.Project.ProjectNode,Microsoft.VisualStudio.Shell.Interop.VSCOMPONENTSELECTORDATA,System.String)")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1703:ResourceStringsShouldBeSpelledCorrectly", MessageId = "pdftex", Scope = "resource", Target = "VsTeXProject.VisualStudio.Project.resources")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1703:ResourceStringsShouldBeSpelledCorrectly", MessageId = "xbb", Scope = "resource", Target = "VsTeXProject.VisualStudio.Project.resources")]
[assembly: SuppressMessage("Microsoft.Naming", "CA1703:ResourceStringsShouldBeSpelledCorrectly", MessageId = "platex", Scope = "resource", Target = "VsTeXProject.VisualStudio.Project.resources")]

