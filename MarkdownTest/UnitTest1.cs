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
norm**al `this is inline_code()`, no**problem
```
var foo = function (bar) {
  return bar++;
};

console.log(foo(5));
```

![Base64 image](data:image/png;base64, iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==)

![Image attachment](https://cdn.pixabay.com/index/2021/08/24/12-14-41-390_1440x550.jpg)

normal text2

new line  
test

second new line 
test
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
