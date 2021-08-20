using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdown2Openxml.Enumeration;

namespace Markdown2Openxml.ParagraphProcessor
{
    public class Heading1ParagraphProcessor : ParagraphProcessorInterface
    {

        public IList<OpenXmlCompositeElement> process(MainDocumentPart mainDocumentPart, StringArrayReader reader)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();

            RunProperties runProperties1 = new RunProperties();
            FontSize fontSize1 = new FontSize() { Val = "42" };
            runProperties1.Append(fontSize1);

            run.Append(runProperties1);

            string input = reader.getCurrentString().Substring(2);
            run.AppendChild(new Text(input));

            paragraph.Append(run);
            return new List<OpenXmlCompositeElement>(new []{ paragraph });
        }
    }
}
