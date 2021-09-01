using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Markdown2Openxml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using System;
using System.Collections.Generic;

namespace MarkdownTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMarkdownString()
        {
            string markdownString = @"
## h2 Heading
### h3 Heading
norm**al `text`, no**problem
```
var foo = function (bar) {
  return bar++;
};

console.log(foo(5));
```
normal text2
";
            string filepath = @"test.docx";
            using (MemoryStream mem = new MemoryStream())
            {
                using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
                {
                    // MainDocumentPart
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    IList<OpenXmlCompositeElement> paragraphs = MarkdownToOpenxmlUtil.markdownToParagraph(mainPart, markdownString);
                    body.Append(paragraphs);

                    // Debug...
                    foreach (var item in paragraphs)
                    {
                        Console.WriteLine(item + " " + item.InnerText);
                    }
                    
                    mainPart.Document.Save();
                }
                File.WriteAllBytes(filepath, mem.ToArray()); 
            }
        }
    }
}
