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

namespace Markdown2Openxml.RunProcessor
{
    public class ProcessRunTextService
    {   
        private Regex linkRegex = new Regex(@"(\[[a-zA-Z0-9\:\/\.\,\?\\\*-_# ]+\])(\([a-zA-Z0-9\:\/\.\,\?\\\*\-_#]+\))");
        private Boolean isInlineCode = false;

        private void addTextToCurrentRun(MainDocumentPart mainDocumentPart, int index, SortedDictionary<int, Run> runs, StringBuilder stringBuilder, RunProperties runProperties){
            if(stringBuilder.Length == 0)return;

            Run run = new Run();

            //Get the string
            string results = stringBuilder.ToString();

            // Check links, if splitLinks more than normal one, there are some links
            string[] splitLinks = linkRegex.Split(results);
            if(!isInlineCode && splitLinks.Length > 1){
                //Link handling
                int splitIndex = index;
                for(int i = 0 ; i < splitLinks.Length; i++){
                    //Skip empty string
                    if(string.IsNullOrEmpty(splitLinks[i]))continue;

                    RunProperties splitProperties = null;
                    if(runProperties != null){
                        splitProperties = (RunProperties)runProperties.CloneNode(true);
                    }else{
                        splitProperties = new RunProperties();
                    }

                    //Create new Run and add properties
                    splitIndex += splitLinks[i].Length;
                    Run linkRun = new Run();
                    linkRun.Append(splitProperties);


                    //Search for [xxx](yyyy) hyperlinks from markdown content
                    if(i != splitLinks.Length - 1 && 
                        splitLinks[i][0] == '[' && splitLinks[i][splitLinks[i].Length - 1] == ']' &&
                        splitLinks[i + 1][0] == '(' && splitLinks[i + 1][splitLinks[i + 1].Length - 1] == ')'
                        ){
                            Text linkText = new Text(splitLinks[i].Substring(1, splitLinks[i].Length - 2));
                            linkText.Space = SpaceProcessingModeValues.Preserve;

                            // Change uri color to blue
                            Color color = new Color() { Val = "015692" };
                            splitProperties.Append(color);

                            //Add Uri to document part
                            string url = splitLinks[i + 1].Substring(1, splitLinks[i + 1].Length - 2);

                            if(url.StartsWith("http://") || url.StartsWith("https://")){
                                string documentId = "URL-"+System.Guid.NewGuid().ToString();
                                mainDocumentPart.AddHyperlinkRelationship(new Uri(url), true, documentId);

                                //Add to run
                                linkRun.Append(
                                    new Hyperlink(new Run(linkText))
                                    {
                                        Id = documentId
                                    }
                                );
                            }else{
                                //Add to run
                                linkRun.Append(
                                    new Hyperlink(new Run(linkText))
                                    {
                                        Anchor = url
                                    }
                                );
                            }

                            // Skip url part
                            i += 1;
                    }else{
                        //Normal text
                        Text text = new Text(splitLinks[i]);
                        text.Space = SpaceProcessingModeValues.Preserve;
                        linkRun.Append(text);
                    }

                    runs.Add(splitIndex, linkRun);
                }
            }else{
                Text text = new Text(results);
                text.Space = SpaceProcessingModeValues.Preserve;

                if(runProperties != null){
                    run.Append(runProperties);
                }
                //Add to run and reset
                run.Append(text);

                runs.Add(index, run);
            }
        }

        public IList<Run> process(MainDocumentPart mainDocumentPart, string msg){
            SortedDictionary<int, Run> runs = new SortedDictionary<int, Run>();

            int italicsStartPos = -1;
            int boldStartPos = -1;

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(""); //Default

            RunProperties runProperties = new RunProperties();
            HashSet<RunPattern> currentStatus = new HashSet<RunPattern>();
            HashSet<RunPattern> previousStatus = new HashSet<RunPattern>();
            /*
             * Handle logic:
             * 
             * Find the starter (S* / *) records, after found ender (*S / *), 
             * reset all settings and process again. 'S' for space.
             * 
             *      | <------------------------------------------------ |
             *     S*    S*    S    S**        S    **S*    **    **    *
             * test *test *test test **testtest test** *test**test**test*test
             * 
             */
            
            for (int i = 0; i < msg.Length; i++){
                char ch = msg[i];
                char nextCh = i == msg.Length - 1 ? '\0' : msg[i + 1];
                char previousCh = i == 0 ? '\0' : msg[i - 1];

                // Handle bold
                if (ch.Equals('*') && nextCh.Equals('*'))
                {
                    if(currentStatus.Contains(RunPattern.Bold)){
                        currentStatus.Remove(RunPattern.Bold);
                        boldStartPos = -1;
                        i++; //Skip *
                        continue;
                    }else{
                        if (boldStartPos == -1)
                        {
                            addTextToCurrentRun(mainDocumentPart, i, runs, stringBuilder, runProperties);
                            stringBuilder = stringBuilder.Clear();
                            runProperties = new RunProperties();

                            boldStartPos = i;
                            i++; //Skip *
                            continue;
                        }
                        else
                        {
                            //Reset the result and process again (since ** is being added at the start)
                            // we need to delete the ** but to delete **, we need to remove the text 
                            // that leads to the start of **
                            List<int> toRemove = runs.Where(p => p.Key > boldStartPos + 1)
                            .Select(p => p.Key)
                            .ToList();
                            foreach (var key in toRemove)
                            {
                                runs.Remove(key);
                            }

                            i = boldStartPos + 1;

                            // Bold is a OnOffType (OpenXML will process it like a switch.)
                            // eg. "<bold> ... <bold>" -> "<bold> ... </bold>" (on ... off)
                            currentStatus.Add(RunPattern.Bold);
                            stringBuilder.Clear();
                            continue;
                        }
                    }
                //handle italic S*
                }else if(italicsStartPos != -1 && ch.Equals('*') && !nextCh.Equals('*') && !previousCh.Equals(' ') && !previousCh.Equals('\0')){
                    if(currentStatus.Contains(RunPattern.Italic)){
                        currentStatus.Remove(RunPattern.Italic);
                        italicsStartPos = -1;
                        continue;
                    }else if ( italicsStartPos != -1 ){
                        //Reset the result and process again
                        List<int> toRemove = runs.Where(p => p.Key > italicsStartPos)
                        .Select(p => p.Key)
                        .ToList();
                        foreach (var key in toRemove)
                        {
                            runs.Remove(key);
                        }

                        i = italicsStartPos;
                        currentStatus.Add(RunPattern.Italic);
                        stringBuilder.Clear();
                        continue;
                    }
                //handle italic *S
                }else if (ch.Equals('*') && !nextCh.Equals('*') && !nextCh.Equals(' ') && !nextCh.Equals('\0')){
                    addTextToCurrentRun(mainDocumentPart, i, runs, stringBuilder, runProperties);
                    stringBuilder = stringBuilder.Clear();
                    runProperties = new RunProperties();

                    italicsStartPos = i;
                }else if(ch == '`'){
                    // self-contained parsing (ie. i += n)
                    // add the stuff behind ` to paragraph
                    addTextToCurrentRun(mainDocumentPart, i, runs, stringBuilder, runProperties);
                    stringBuilder = stringBuilder.Clear();
                    runProperties = new RunProperties();
                    previousStatus.Clear();

                    // look ahead (bypasses "Styling" part)
                    isInlineCode = true;
                    int pairLocation = msg.IndexOf('`', i+1);
                    if(pairLocation != -1 && msg[pairLocation-1] != '\\'){
                        string inlineCodeText = msg.Substring(i+1, pairLocation-i-1);
                        RunProperties inlineCodeStyle = new RunProperties();
                        inlineCodeStyle.Append(new FontSize(){ Val = "20" });
                        inlineCodeStyle.Append(new Color() { Val = "808080" });
                        addTextToCurrentRun(mainDocumentPart, i+1, runs, stringBuilder.Append(inlineCodeText), inlineCodeStyle);
                        i += pairLocation - i;
                        stringBuilder.Clear();
                        continue;
                    }
                }

                //Styling
                if(!previousStatus.SetEquals(currentStatus)){
                    addTextToCurrentRun(mainDocumentPart, i, runs, stringBuilder, runProperties);
                    stringBuilder = stringBuilder.Clear();
                    runProperties = new RunProperties();

                    if (currentStatus.Count > 0) {

                        //Bold handling
                        if (currentStatus.Contains(RunPattern.Bold))
                        {
                            Bold bold = new Bold();
                            bold.Val = OnOffValue.FromBoolean(true);
                            runProperties.Append(bold);

                        }

                        //Italic handling
                        if (currentStatus.Contains(RunPattern.Italic))
                        {
                            Italic italic = new Italic();
                            italic.Val = OnOffValue.FromBoolean(true);
                            runProperties.Append(italic);
                        }
                    }
                }

                previousStatus = new HashSet<RunPattern>(currentStatus);
                stringBuilder.Append(ch);
            }
            //Ending paragraph
            addTextToCurrentRun(mainDocumentPart, msg.Length, runs, stringBuilder, runProperties);

            return runs.Select(p => p.Value).ToList();
        }

        
    }
}
