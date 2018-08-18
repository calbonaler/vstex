using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{
	[Export(typeof(IClassifierProvider))]
	[ContentType("vstex")]
	internal class TeXClassifierProvider : IClassifierProvider
	{
		/// <summary>
		///     Import the classification registry to be used for getting a reference
		///     to the custom classification type later.
		/// </summary>
		[Import]
		internal IClassificationTypeRegistryService classificationRegistry = null;

		[Import]
		internal IBufferTagAggregatorFactoryService tagAggregatorFactory = null;

		public TeXClassifierProvider() { }

		public IClassifier GetClassifier(ITextBuffer buffer) => buffer.Properties.GetOrCreateSingletonProperty(() => new TeXClassifier(buffer, classificationRegistry, tagAggregatorFactory));
	}

	internal class TeXClassifier : IClassifier
	{
		private readonly IClassificationType _commentType;
		private readonly IClassificationType _beginEndType;
		private readonly IClassificationType _functionType;
		private readonly IClassificationType _braceType;

		internal TeXClassifier(ITextBuffer buffer, IClassificationTypeRegistryService registry, IBufferTagAggregatorFactoryService factory)
		{
			_commentType = registry.GetClassificationType(nameof(TeXClassifierFormat.TeXCommentFormat));
			_beginEndType = registry.GetClassificationType(nameof(TeXClassifierFormat.TeXEnvironmentHeaderFormat));
			_functionType = registry.GetClassificationType(nameof(TeXClassifierFormat.TeXFunctionFormat));
			_braceType = registry.GetClassificationType(nameof(TeXClassifierFormat.TeXDelimiterFormat));
		}

		public IList<ClassificationSpan> GetClassificationSpans(SnapshotSpan span) => GetClassificationSpansInternal(span).ToArray();

		IEnumerable<ClassificationSpan> GetClassificationSpansInternal(SnapshotSpan span)
		{
			var text = span.GetText();
			for (var pt = 0; pt < span.Length;)
			{
				if (text[pt] == '\\')
				{
					if (pt + 1 >= span.Length)
					{
						yield return new ClassificationSpan(new SnapshotSpan(span.Start + pt, 1), _functionType);
						pt += 1;
						continue;
					}
					var isBegin = ContainsExact(text, pt + 1, "begin");
					if (isBegin || ContainsExact(text, pt + 1, "end"))
					{
						var end = pt + (isBegin ? 6 : 4);
						while (end < span.Length && char.IsWhiteSpace(text[end])) end++;
						if (end < span.Length && text[end] == '{')
						{
							while (end < span.Length && text[end] != '}') end++;
							if (end < span.Length && text[end] == '}')
							{
								end++;
								yield return new ClassificationSpan(new SnapshotSpan(span.Start + pt, end - pt), _beginEndType);
								pt = end;
								continue;
							}
						}
					}
					if (text[pt + 1] >= 'A' && text[pt + 1] <= 'Z' || text[pt + 1] >= 'a' && text[pt + 1] <= 'z' || text[pt + 1] == '@')
					{
						var end = pt + 2;
						while (end < span.Length && (text[end] >= 'A' && text[end] <= 'Z' || text[end] >= 'a' && text[end] <= 'z' || text[end] == '@'))
							end++;
						yield return new ClassificationSpan(new SnapshotSpan(span.Start + pt, end - pt), _functionType);
						pt = end;
						continue;
					}
					yield return new ClassificationSpan(new SnapshotSpan(span.Start + pt, 2), _functionType);
					pt += 2;
					continue;
				}
				if (text[pt] == '%')
				{
					yield return new ClassificationSpan(new SnapshotSpan(span.Start + pt, text.Length - pt), _commentType);
					break;
				}
				if (text[pt] == '{' || text[pt] == '}' || text[pt] == '[' || text[pt] == ']' || text[pt] == '$' || text[pt] == '&')
				{
					yield return new ClassificationSpan(new SnapshotSpan(span.Start + pt, 1), _braceType);
					pt += 1;
					continue;
				}
				pt += 1;
			}
		}

		static bool ContainsExact(string baseString, int offset, string subString)
		{
			if (offset < 0 || baseString.Length < offset + subString.Length)
				return false;
			for (var i = 0; i < subString.Length; i++)
			{
				if (baseString[i + offset] != subString[i])
					return false;
			}
			return true;
		}

		/// <summary>
		/// Create an event for when the Classification changes
		/// </summary>
#pragma warning disable 67
		public event EventHandler<ClassificationChangedEventArgs> ClassificationChanged;
#pragma warning restore 67
	}
}
