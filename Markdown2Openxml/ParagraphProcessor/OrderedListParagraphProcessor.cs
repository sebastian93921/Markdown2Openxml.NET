using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdown2Openxml.Enumeration;
using Markdown2Openxml.RunProcessor;

namespace Markdown2Openxml.ParagraphProcessor
{
    public class OrderedListParagraphProcessor : ParagraphProcessorInterface
    {

        private static ProcessRunTextService processRunTextService = new ProcessRunTextService();

        public IList<OpenXmlCompositeElement> process(MainDocumentPart mainDocumentPart, StringArrayReader reader)
        {
            Regex unorderListRegex = MarkdownPatternProcessor.ParagraphPatterns[ParagraphPattern.OrderedList];
            string[] inputArray = unorderListRegex.Split(reader.getCurrentString());
            
            string inputString = "";
            if(inputArray.Length > 1){
                inputString = MarkdownToOpenxmlUtil.ParagraphLineConcat(inputArray[1], reader);
            }

            Paragraph element =
                new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId() { Val = "ListParagraph" },
                        new NumberingProperties(
                            new NumberingLevelReference() { Val = 0 },
                            new NumberingId() { Val = 5 }
                        )
                    )
                ) { RsidParagraphAddition = "00031711", RsidParagraphProperties = "00031711", RsidRunAdditionDefault = "00031711" };
            element.Append(processRunTextService.process(mainDocumentPart, inputString));

            return new List<OpenXmlCompositeElement>(){ element };
        }
    }
}
