using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdown2Openxml.Enumeration;
using System.Text.RegularExpressions;
using Markdown2Openxml.RunProcessor;

namespace Markdown2Openxml.ParagraphProcessor
{
    public class TableParagraphProcessor : ParagraphProcessorInterface
    {

        private ProcessRunTextService processRunTextService = new ProcessRunTextService();

        public IList<OpenXmlCompositeElement> process(MainDocumentPart mainDocumentPart, StringArrayReader reader)
        {
            Table table = new Table();
            table.AppendChild((TableProperties)MarkdownToOpenxmlUtil.commonTableProperties.CloneNode(true));

            while (!reader.endOfLine())
            {
                string line = reader.getCurrentString();

                string[] values = line.Split("|").Skip(1).SkipLast(1).ToArray();

                if(values.Length > 0){
                    TableRow tr = new TableRow();

                    bool isHeader = checkIsHeader(reader.nextLineString()); 

                    if(isHeader){
                        reader.increasePos();
                    }

                    foreach(string value in values){
                        TableCell tc = new TableCell();

                        if(isHeader){
                            Run headerRun = new Run();

                            RunProperties runProperties = new RunProperties();
                            Bold bold = new Bold();
                            bold.Val = OnOffValue.FromBoolean(true);             
                            runProperties.Append(bold);

                            headerRun.Append(runProperties);
                            headerRun.Append(new Text(value));

                            tc.Append(new Paragraph(headerRun));
                            tr.Append(tc);
                        }else{
                            tc.Append(new Paragraph(processRunTextService.process(mainDocumentPart, value)));
                            tr.Append(tc);
                        }
                    }

                    table.Append(tr);
                }

                if (MarkdownPatternProcessor.getParagraphPattern(reader.nextLineString()) != ParagraphPattern.Table){
                    return new List<OpenXmlCompositeElement>(new[] { new Paragraph(new Run(table)) }); ;
                }

                //Increase
                reader.increasePos();
            }
            return new List<OpenXmlCompositeElement>();
        }

        private bool checkIsHeader(string nextLine){
            if(nextLine == null)return false;

            Regex headerCheck = new Regex(@"(\| ?(---*) ?)+\|");
            if (headerCheck.IsMatch(nextLine)){
                return true;
            }

            return false;
        }
    }
}
