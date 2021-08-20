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
using Markdown2Openxml.RunProcessor;
using Markdown2Openxml.Model;
using System.Collections;

namespace Markdown2Openxml
{
    public class SimpleSyntaxHighlightUtil
    {
        private static readonly string COMMENT_COLOR = "aaaaaa";
        private static readonly string FUNCTION_COLOR = "c540c3";
        private static readonly string STRING_COLOR = "128953";
        private static readonly string DATATYPE_COLOR = "ff0603";
        private static readonly string DEFAULT_COLOR = "282828";
        private static readonly string FUNCTION_NAME_COLOR = "4078f2";

        public static IList<Run> ParselineToRuns(string line){
            IList<Run> runs = new List<Run>();

            IList<ColorStyle> styleList = ParselineToStyleList(line);
            foreach (ColorStyle token in styleList)   
            {   
                // Create a run class for each words
                Run run = new Run();
                RunProperties runPropertiesCode = new RunProperties();

                runPropertiesCode.Append(new FontSize(){ Val = "18" });
                runPropertiesCode.Append(new Color() { Val = token.Color });
                run.Append(runPropertiesCode);
                Text codeLine = new Text(token.Description);
                codeLine.Space = SpaceProcessingModeValues.Preserve;
                run.Append(codeLine);

                runs.Add(run); 
            }

            return runs;
        }

        public static IList<ColorStyle> ParselineToStyleList(string line){
            IList<ColorStyle> styleList = new List<ColorStyle>();

            Regex splitCodeRegex = new Regex("([ \\t{}():;.])");  
            Regex functionNameRegex = new Regex(@"[\w]+\(");
            String [] tokens = splitCodeRegex.Split(line);   
            bool isComment = false;
            for (int pos = 0; pos < tokens.Length; pos++)   
            {
                string token = tokens[pos];
                // Default color
                string color = DEFAULT_COLOR; 

                if(!isComment){
                    // Function name check
                    bool isFunctionName = false;
                    if (pos < tokens.Length - 1 && functionNameRegex.IsMatch(token + tokens[pos+1])){
                        for(int j = pos ; j < tokens.Length; j++){
                            if(tokens[j].Equals(")")){
                                isFunctionName = true;
                            }
                        }
                    }

                    if(isFunctionName){
                        //Function name
                        color = FUNCTION_NAME_COLOR;
                    }
                    // Check comment
                    else if (token.StartsWith("//")){
                        isComment = true;
                        color = COMMENT_COLOR;
                    }
                    // Check whether the token is a keyword. 
                    else if (Regex.IsMatch(token, "^['\"](.*)['\"]$")){
                        //String value
                        color = STRING_COLOR;
                    }else if (token.StartsWith("@")){
                        //@Annotation
                        color = DATATYPE_COLOR;
                    }else{
                        //Functional keywords
                        String [] functionalKeywords = { "abstract", "continue", "for", "new",
                                        "switch", "assert", "default", "goto",
                                        "package", "synchronized", "boolean",
                                        "do", "if", "private", "this", "break",
                                        "double", "implements", "protected",
                                        "throw", "byte", "else", "import",
                                        "public", "throws", "case", "enum",
                                        "instanceof", "return", "transient",
                                        "catch", "extends", "int", "short",
                                        "try", "char", "final", "interface",
                                        "static", "void", "class", "finally",
                                        "long", "strictfp", "volatile", "const",
                                        "float", "native", "super", "while", "string" };
                        for (int i = 0; i < functionalKeywords.Length; i++)  
                        {  
                            if (functionalKeywords[i] == token)  
                            {  
                                color = FUNCTION_COLOR; 
                                break;  
                            }
                        }
                    }
                }else{
                    color = COMMENT_COLOR;
                }

                styleList.Add(new ColorStyle(token, color)); 
            }

            return styleList;
        }

        public static string ParselineToHtml(string line){
            StringBuilder sb = new StringBuilder();
            IList<ColorStyle> styleList = ParselineToStyleList(line);

            foreach (ColorStyle token in styleList)   
            {   
                // create a code of block for each words
                sb.Append($"<span style=\"color: #{token.Color}\">{token.Description}</span>");
            }

            return sb.ToString();
        }

    }
}
