using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Markdown2Openxml.Enumeration;

namespace Markdown2Openxml
{
    public class MarkdownPatternProcessor
    {

        public static Dictionary<ParagraphPattern, Regex> ParagraphPatterns = new Dictionary<ParagraphPattern, Regex>()
		{
			{ ParagraphPattern.CodeBlock, new Regex(@"^```(.*)") },
			{ ParagraphPattern.HorizontalRule, new Regex(@"^\* \* \*( \*)*$") },
			{ ParagraphPattern.Heading1, new Regex(@"^# (.*)") },
			{ ParagraphPattern.Heading2, new Regex(@"^## (.*)") },
			{ ParagraphPattern.Heading3, new Regex(@"^### (.*)") },
			{ ParagraphPattern.Image, new Regex(@"^\!\[(.+?)\]\((.+)\)") },
			{ ParagraphPattern.OrderedList, new Regex(@"^[\d]\. (.*)") },
			{ ParagraphPattern.Quote, new Regex(@"^>{1} (.*)") },
			{ ParagraphPattern.Table, new Regex(@"^\|(.*)\|") },
			{ ParagraphPattern.UnorderedList, new Regex(@"^[*+-] (.*)") },
			{ ParagraphPattern.None, new Regex("(.*)") }
		};

        public static ParagraphPattern getParagraphPattern(string markdown)
		{
			if(markdown == null) return ParagraphPattern.None;

			foreach (var pattern in ParagraphPatterns)
			{
				var regex = pattern.Value;
				if (!regex.IsMatch(markdown)) continue;
				return pattern.Key;
			}
            return ParagraphPattern.None;
		}

    }
}
