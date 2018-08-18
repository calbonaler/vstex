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
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.VisualStudio.Shell.Interop;

namespace VsTeXProject.VisualStudio.Project
{
    public class BuildActionConverter : EnumConverter
    {
        public BuildActionConverter() : base(typeof (BuildAction)) { }

		public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType) => sourceType == typeof(string) || base.CanConvertFrom(context, sourceType);

		public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            var str = value as string;
            if (str != null)
            {
                if (str == SR.GetString(SR.Compile, culture)) return BuildAction.Compile;
                if (str == SR.GetString(SR.Picture, culture)) return BuildAction.Picture;
                if (str == SR.GetString(SR.Content, culture)) return BuildAction.Content;
            }
            return BuildAction.Content;
        }

        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof (string))
            {
                // In some cases if multiple nodes are selected the windows form engine calls us with a null value if the selected node's property values are not equal
                // Example of windows form engine passing us null: File set to Compile, Another file set to None, bot nodes are selected, and the build action combo is clicked.
                return SR.GetString((value != null ? (BuildAction)value : BuildAction.Content).ToString(), culture);
            }
            return base.ConvertTo(context, culture, value, destinationType);
        }

		public override bool GetStandardValuesSupported(ITypeDescriptorContext context) => true;

		public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext context) => new StandardValuesCollection(new[] { BuildAction.Compile, BuildAction.Picture, BuildAction.Content });
	}

	public class TeXProcessorConverter : TypeConverter
	{
		readonly string[] standardValues = new string[] { "platex", "uplatex", "latex", "pdflatex", "lualatex", "xelatex" };

		public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType) => sourceType == typeof(string) || base.CanConvertFrom(context, sourceType);

		public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value) => value is string str ? value : base.ConvertFrom(context, culture, value);

		public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType) => destinationType == typeof(string) || base.CanConvertTo(context, destinationType);

		public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType) => destinationType == typeof(string) ? ((string)value ?? standardValues[0]) : base.ConvertTo(context, culture, value, destinationType);

		public override bool GetStandardValuesSupported(ITypeDescriptorContext context) => true;

		public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext context) => new StandardValuesCollection(standardValues);
	}
}