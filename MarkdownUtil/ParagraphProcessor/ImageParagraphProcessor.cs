using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdown2Openxml.Enumeration;

namespace Markdown2Openxml.ParagraphProcessor
{
    public class ImageParagraphProcessor : ParagraphProcessorInterface
    {

        public IList<OpenXmlCompositeElement> process(MainDocumentPart mainDocumentPart, StringArrayReader reader)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();

            Drawing drawing = MarkdownImageProcessor.convertMarkdownImageToRunElement(mainDocumentPart, reader.getCurrentString());
            if (drawing != null)
            {
                run.AppendChild(drawing);
                paragraph.Append(run);
            };

            return new List<OpenXmlCompositeElement>(){ paragraph };
        }
    }
}
