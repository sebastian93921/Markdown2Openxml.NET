using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Http;
using System.Net;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using Markdown2Openxml.Enumeration;
using Markdown2Openxml.ParagraphProcessor;
using Markdown2Openxml.RunProcessor;

namespace Markdown2Openxml
{
    public class MarkdownToOpenxmlUtil
    {
        public static TableProperties commonTableProperties = new TableProperties(
            new TableBorders(
                new TopBorder() { Val = 
                    new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                new BottomBorder() { Val = 
                    new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                new LeftBorder() { Val = 
                    new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                new RightBorder() { Val = 
                    new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                new InsideHorizontalBorder() { Val = 
                    new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 },
                new InsideVerticalBorder() { Val = 
                    new EnumValue<BorderValues>(BorderValues.Thick), Size = 6 }
            ),
            new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableLayout(){ Type = TableLayoutValues.Fixed },
            new TableCellMarginDefault(
                new TopMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
                new StartMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
                new BottomMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
                new EndMargin() { Width = "100", Type = TableWidthUnitValues.Dxa }
            )
        );

        public static Dictionary<ParagraphPattern, ParagraphProcessorInterface> ParagraphProcessors = new Dictionary<ParagraphPattern, ParagraphProcessorInterface>()
		{
			{ ParagraphPattern.CodeBlock, new CodeBlockParagraphProcessor() },
			{ ParagraphPattern.Heading1, new Heading1ParagraphProcessor() },
			{ ParagraphPattern.Heading2, new Heading2ParagraphProcessor() },
			{ ParagraphPattern.Heading3, new Heading3ParagraphProcessor() },
			{ ParagraphPattern.Image, new ImageParagraphProcessor() },
            { ParagraphPattern.Table, new TableParagraphProcessor() },
            { ParagraphPattern.UnorderedList, new UnorderListParagraphProcessor() },
            { ParagraphPattern.OrderedList, new OrderedListParagraphProcessor() },
            { ParagraphPattern.HorizontalRule, new HorizontalRuleParagraphProcessor() }
		};

        public static ProcessRunTextService processRunTextService = new ProcessRunTextService();

        public MarkdownToOpenxmlUtil()
        {
        }

        public static IList<OpenXmlCompositeElement> markdownToParagraph(MainDocumentPart mainDocumentPart, string markdownString)
        {
            IList<OpenXmlCompositeElement> paragraphs = new List<OpenXmlCompositeElement>();

            // Null check
            if(String.IsNullOrEmpty(markdownString)){
                paragraphs.Add(new Paragraph());
                return paragraphs;
            }

            StringArrayReader stringArrayReader = new StringArrayReader(markdownString.Split(
                new[] { Environment.NewLine },
                StringSplitOptions.None
            ));

            while (!stringArrayReader.endOfLine())
            {
                string line = stringArrayReader.getCurrentString();

                ParagraphPattern pattern = MarkdownPatternProcessor.getParagraphPattern(line);

                if (ParagraphProcessors.ContainsKey(pattern))
                {
                    ParagraphProcessorInterface paragraphProcessor = ParagraphProcessors[pattern];
                    ((List<OpenXmlCompositeElement>)paragraphs).AddRange(paragraphProcessor.process(mainDocumentPart, stringArrayReader));
                }
                else
                {
                    string paragraphString = ParagraphLineConcat(line, stringArrayReader);
                    Paragraph defaultPara = new Paragraph(processRunTextService.process(mainDocumentPart, paragraphString));
                    paragraphs.Add(defaultPara);
                }

                stringArrayReader.increasePos();
            }

            // If no paragraph, add an empty one
            if (paragraphs.Count == 0)paragraphs.Add(new Paragraph());
            return paragraphs;
        }

        public static string ParagraphLineConcat(string firstLine, StringArrayReader stringArrayReader){
            // Normal String handling
            string paragraphString = firstLine;
            if(!String.IsNullOrEmpty(firstLine)){
                while (!stringArrayReader.endOfLine())
                {
                    // Check if the next line of string is the same paragraphs or not
                    string nextLineString = stringArrayReader.nextLineString();
                    if (nextLineString != null && 
                        MarkdownPatternProcessor.getParagraphPattern(nextLineString) == ParagraphPattern.None && 
                        !String.IsNullOrWhiteSpace(nextLineString.Trim()))
                    {
                        paragraphString += nextLineString;
                    }else{
                        break;
                    }

                    //Increase
                    stringArrayReader.increasePos();
                }
            }
            return paragraphString;
        }

    }
}
