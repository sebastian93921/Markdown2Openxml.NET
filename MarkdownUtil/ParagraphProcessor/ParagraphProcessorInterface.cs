using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Markdown2Openxml.ParagraphProcessor
{
    public interface ParagraphProcessorInterface
    {
        IList<OpenXmlCompositeElement> process(MainDocumentPart mainDocumentPart, StringArrayReader reader);
    }
}
