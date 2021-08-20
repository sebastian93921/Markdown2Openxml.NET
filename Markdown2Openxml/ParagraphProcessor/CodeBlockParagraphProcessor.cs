using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdown2Openxml.Enumeration;

namespace Markdown2Openxml.ParagraphProcessor
{
    public class CodeBlockParagraphProcessor : ParagraphProcessorInterface
    {

        public IList<OpenXmlCompositeElement> process(MainDocumentPart mainDocumentPart, StringArrayReader reader)
        {
            Table codeTable = new Table();
            codeTable.AppendChild((TableProperties)MarkdownToOpenxmlUtil.commonTableProperties.CloneNode(true));

            TableRow tr = new TableRow();
            TableCell tc = new TableCell();
            TableCellProperties tcp = new TableCellProperties(
                // Cell Background color
                new DocumentFormat.OpenXml.Wordprocessing.Shading() {
                    Color = "auto",
                    Fill = "f6f6f6",
                    Val = ShadingPatternValues.Clear
                }
            );
            tc.Append(tcp);

            IList<Paragraph> codeParas = new List<Paragraph>();

            reader.increasePos();
            while (!reader.endOfLine())
            {
                string line = reader.getCurrentString();
                if (MarkdownPatternProcessor.getParagraphPattern(line) == ParagraphPattern.CodeBlock)
                {
                    tc.Append(codeParas);
                    tr.Append(tc);
                    codeTable.Append(tr);

                    return new List<OpenXmlCompositeElement>(new[] { new Paragraph(new Run(codeTable)) }); ;
                }

                codeParas.Add(new Paragraph(SimpleSyntaxHighlightUtil.ParselineToRuns(line)));

                //Increase
                reader.increasePos();
            }
            return new List<OpenXmlCompositeElement>();
        }
    }
}
