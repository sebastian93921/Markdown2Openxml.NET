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
    public class HorizontalRuleParagraphProcessor : ParagraphProcessorInterface
    {
        public IList<OpenXmlCompositeElement> process(MainDocumentPart mainDocumentPart, StringArrayReader reader)
        {
            Paragraph element = new Paragraph();
            ParagraphProperties paraProperties = new ParagraphProperties();
            ParagraphBorders paraBorders = new ParagraphBorders();
            BottomBorder bottom = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)12U, Space = (UInt32Value)1U };
            paraBorders.Append(bottom);
            paraProperties.Append(paraBorders);
            element.Append(paraProperties);

            return new List<OpenXmlCompositeElement>(){ element };
        }
    }
}
