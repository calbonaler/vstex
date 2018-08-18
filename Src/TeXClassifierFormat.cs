using System.ComponentModel.Composition;
using System.Windows.Media;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{
	static class TeXClassifierType
    {
		[Export(typeof(ClassificationTypeDefinition))]
		[Name(nameof(TeXClassifierFormat.TeXCommentFormat))]
		internal static ClassificationTypeDefinition TeXCommentFormat = null;

        [Export(typeof(ClassificationTypeDefinition))]
		[Name(nameof(TeXClassifierFormat.TeXEnvironmentHeaderFormat))]
		internal static ClassificationTypeDefinition TeXEnvironmentHeaderFormat = null;

        [Export(typeof(ClassificationTypeDefinition))]
		[Name(nameof(TeXClassifierFormat.TeXFunctionFormat))]
		internal static ClassificationTypeDefinition TeXFunctionFormat = null;

        [Export(typeof(ClassificationTypeDefinition))]
		[Name(nameof(TeXClassifierFormat.TeXDelimiterFormat))]
		internal static ClassificationTypeDefinition TeXDelimiterFormat = null;
    }

    static class TeXClassifierFormat
    {
        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = nameof(TeXCommentFormat))]
        [Name(nameof(TeXCommentFormat))]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.Low)]
        [Order(After = Priority.Default)]
        internal sealed class TeXCommentFormat : ClassificationFormatDefinition
        {
            public TeXCommentFormat()
            {
                DisplayName = "TeX - Comment"; //human readable version of the name
				ForegroundColor = Colors.Red;
            }
        }

        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = nameof(TeXEnvironmentHeaderFormat))]
        [Name(nameof(TeXEnvironmentHeaderFormat))]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.Low)]
        [Order(After = Priority.High)]
        internal sealed class TeXEnvironmentHeaderFormat : ClassificationFormatDefinition
        {
            public TeXEnvironmentHeaderFormat()
            {
                DisplayName = "TeX - Environment Header"; //human readable version of the name
				ForegroundColor = Colors.Green;
            }
        }

        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = nameof(TeXFunctionFormat))]
        [Name(nameof(TeXFunctionFormat))]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.Low)]
        [Order(After = Priority.Default)]
        internal sealed class TeXFunctionFormat : ClassificationFormatDefinition
        {
            public TeXFunctionFormat()
            {
                DisplayName = "TeX - Function"; //human readable version of the name
				ForegroundColor = Colors.Blue;
            }
        }

        [Export(typeof(EditorFormatDefinition))]
        [ClassificationType(ClassificationTypeNames = nameof(TeXDelimiterFormat))]
        [Name(nameof(TeXDelimiterFormat))]
        [UserVisible(true)] //this should be visible to the end user
        [Order(Before = Priority.Default)]
        [Order(After = Priority.Default)]
        internal sealed class TeXDelimiterFormat : ClassificationFormatDefinition
        {
            public TeXDelimiterFormat()
            {
                DisplayName = "TeX - Delimiter"; //human readable version of the name
				ForegroundColor = Color.FromRgb(163, 21, 21);
            }
        }
    }
}
