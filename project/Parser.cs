using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication4
{
    public class CssParser
    {
        private List<string> _styleSheets;
        private SortedList<string, StyleClass> _scc;
        public SortedList<string, StyleClass> Styles
        {
            get { return this._scc; }
            set { this._scc = value; }
        }

        public CssParser()
        {
            this._styleSheets = new List<string>();
            this._scc = new SortedList<string, StyleClass>();
        }

        public void AddStyleSheet(string path)
        {
            this._styleSheets.Add(path);
            ProcessStyleSheet(path);
        }

        public string GetStyleSheet(int index)
        {
            return this._styleSheets[index];
        }

        private void ProcessStyleSheet(string path)
        {
            string content = CleanUp(File.ReadAllText(path));
            string[] parts = content.Split('}');
            foreach (string s in parts)
            {
                if (CleanUp(s).IndexOf('{') > -1)
                {
                    FillStyleClass(s);
                }
            }
        }

        private void FillStyleClass(string s)
        {
            StyleClass sc = null;
            string[] parts = s.Split('{');
            string styleName = CleanUp(parts[0]).Trim().ToLower();

            if (this._scc.ContainsKey(styleName))
            {
                sc = this._scc[styleName];
                this._scc.Remove(styleName);
            }
            else
            {
                sc = new StyleClass();
            }

            sc.Name = styleName;

            string[] atrs = CleanUp(parts[1]).Replace("}", "").Split(';');
            foreach (string a in atrs)
            {
                if (a.Contains(":"))
                {
                    string _key = a.Split(':')[0].Trim().ToLower();
                    if (sc.Attributes.ContainsKey(_key))
                    {
                        sc.Attributes.Remove(_key);
                    }
                    sc.Attributes.Add(_key, a.Split(':')[1].Trim().ToLower());
                }
            }
            this._scc.Add(sc.Name, sc);
        }

        private string CleanUp(string s)
        {
            string temp = s;
            string reg = "(/\\*(.|[\r\n])*?\\*/)|(//.*)";
            Regex r = new Regex(reg);
            temp = r.Replace(temp, "");
            temp = temp.Replace("\r", "").Replace("\n", "");
            return temp;
        }
        public string CleanHtml(string html, string HTMLPath)
        {

            string urlPath = HTMLPath;
            CssParser oParse = new CssParser();
            oParse.AddStyleSheet(urlPath);
            SortedList<string, CssParser.StyleClass> oList = oParse.Styles;
            string BoldSpans = "";
            string ItalicSpans = "";
            string SuperSpans = "";
            string SubSpans = "";
            foreach (string oKey in oList.Keys)
            {
                CssParser.StyleClass oClass = oList[oKey];
                SortedList<string, string> oAttributes = oClass.Attributes;
                foreach (string atr in oAttributes.Keys)
                {
                    if (atr == "font-weight" && oAttributes[atr] == "bold")
                    {
                        BoldSpans = BoldSpans + "||" + oKey + "||";
                    }
                    else if (atr == "font-style" && oAttributes[atr] == "italic")
                    {
                        ItalicSpans = ItalicSpans + "||" + oKey + "||";
                    }
                    else if (atr == "vertical-align" && oAttributes[atr] == "sub")
                    {
                        SubSpans = SubSpans + "||" + oKey + "||";
                    }
                    else if (atr == "vertical-align" && oAttributes[atr] == "super")
                    {
                        SuperSpans = SuperSpans + "||" + oKey + "||";
                    }
                }
            }
            //Remove Span tags and overwrite the font-overrides
            html = cleanSpanTags(html, BoldSpans, ItalicSpans, SuperSpans, SubSpans);
            //Ends here

            //Clear Empty Paragraphs
            html = ClearEmptyParagraphs(html);
            //Ends here

            
            
            Regex reg = new Regex(@"<[/]?(font|span|xml|del|ins|[ovwxp]:\w+)[^>]*?>", RegexOptions.IgnoreCase);
            MatchCollection spanMatches = reg.Matches(html);
            foreach (Match oMat in spanMatches)
            {
                if (oMat.Value.ToLower().Contains("<p ") == false)
                    html = Regex.Replace(html, oMat.Value, "", RegexOptions.IgnoreCase);
            }


            

            reg = new Regex(@"<([^>]*)(?:class|lang|style|size|face|[ovwxp]:\w+)=(?:'[^']*'|""[^""]*""|[^\s>]+)([^>]*)>", RegexOptions.IgnoreCase);
            spanMatches = reg.Matches(html);
            foreach (Match oMat in spanMatches)
            {
                if (oMat.Value.ToLower().Contains("<p ") == false)
                {
                    string tmp = oMat.Value;
                    tmp = reg.Replace(tmp, "<$1$2>");
            
                    html = html.Replace(oMat.Value, tmp);
                }
            }

            
            reg = new Regex(@"<([^>]*)(?:class|lang|style|size|face|[ovwxp]:\w+)=(?:'[^']*'|""[^""]*""|[^\s>]+)([^>]*)>", RegexOptions.IgnoreCase);
            spanMatches = reg.Matches(html);
            foreach (Match oMat in spanMatches)
            {
                if (oMat.Value.ToLower().Contains("<p ") == false)
                {
                    string tmp = oMat.Value;
                    tmp = reg.Replace(tmp, "<$1$2>");
               
                    html = html.Replace(oMat.Value, tmp);
                }
            }
            //Ends here
            return html;
        }
        private string ClearEmptyParagraphs(string inHTML)
        {
            
            Regex paraRegex = new Regex(@"\<p\s*(.*?)\>(&nbsp;|\s)+\<\/p\>", RegexOptions.IgnoreCase);
            inHTML = paraRegex.Replace(inHTML, "");
            paraRegex = new Regex(@"(\<p\s*class=\s*\W*)(MsoNormal_)(\w+\W*\>)", RegexOptions.IgnoreCase);            
            inHTML = paraRegex.Replace(inHTML, "$1$3");
            paraRegex = new Regex(@"(\<p\s*class=\s*\W*)(\w+_)(\w+\W*\>)", RegexOptions.IgnoreCase);
            inHTML = paraRegex.Replace(inHTML, "$1$3");
            return inHTML;
        }
        private string cleanSpanTags(string inHTML, string inBold, string inItalic, string inSuper, string inSub)
        {
            Regex SpanRegex = new Regex(@"(\<span\s*([^<]*)\>)+(.*?)(\<\/span\>)+", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            MatchCollection spanMatches = SpanRegex.Matches(inHTML);
            foreach (Match oMat in spanMatches)
            {
                Regex classRegex = new Regex(@"class\s*=\s*\W+(\w+)\W+", RegexOptions.IgnoreCase);
                string TagApplied = "";
                if (classRegex.IsMatch(oMat.Value))
                {
                    Match clasMatch = classRegex.Matches(oMat.Value)[0];
                    TagApplied = clasMatch.Groups[1].Value;
                }
                if (TagApplied != "")
                {
                    if ((inBold.ToLower().Contains("||" + "span." + TagApplied.ToLower() + "||") || (inBold.ToLower().Contains("||" + TagApplied.ToLower() + "||"))))
                    {
                        inHTML = inHTML.Replace(oMat.Value, @"<b>" + oMat.Value + @"</b>");
                    }
                    if ((inItalic.ToLower().Contains("||" + "span." + TagApplied.ToLower() + "||") || (inItalic.ToLower().Contains("||" + TagApplied.ToLower() + "||"))))
                    {
                        inHTML = inHTML.Replace(oMat.Value, @"<i>" + oMat.Value + @"</i>");
                    }
                    if ((inSuper.ToLower().Contains("||" + "span." + TagApplied.ToLower() + "||") || (inSuper.ToLower().Contains("||" + TagApplied.ToLower() + "||"))))
                    {
                        inHTML = inHTML.Replace(oMat.Value, @"<sup>" + oMat.Value + @"</sup>");
                    }
                    if ((inSub.ToLower().Contains("||" + "span." + TagApplied.ToLower() + "||") || (inSub.ToLower().Contains("||" + TagApplied.ToLower() + "||"))))
                    {
                        inHTML = inHTML.Replace(oMat.Value, @"<sub>" + oMat.Value + @"</sub>");
                    }
                }
                if (oMat.Value.ToLower().Contains("font-weight:bold"))
                {
                    inHTML = inHTML.Replace(oMat.Value, @"<b>" + oMat.Value + @"</b>");
                }
                if (oMat.Value.ToLower().Contains("font-style:italic"))
                {
                    inHTML = inHTML.Replace(oMat.Value, @"<i>" + oMat.Value + @"</i>");
                }
                if (oMat.Value.ToLower().Contains("vertical-align:sub"))
                {
                    inHTML = inHTML.Replace(oMat.Value, @"<sub>" + oMat.Value + @"</sub>");
                }
                if (oMat.Value.ToLower().Contains("vertical-align:super"))
                {
                    inHTML = inHTML.Replace(oMat.Value, @"<sup>" + oMat.Value + @"</sup>");
                }
            }
            return inHTML;
        }
        public string addStylesheetLinking(string styleSheetname, string titleText, string inHTML)
        {
           
            string titleTxt = @"<title>" +  titleText + @"</title>" + System.Environment.NewLine.ToString();
            string linkStyle = "<link rel='stylesheet' type='text/css' href='" + styleSheetname + "'/>" + System.Environment.NewLine.ToString();
            string linktemplate = "<link rel='stylesheet' type='application/vnd.adobe-page-template+xml' href='page-template.xpgt'/>" + System.Environment.NewLine.ToString();
            inHTML= Regex.Replace(inHTML,@"\<Head(.*?)\>(.*?)\<\/Head\>","<head>" +  titleTxt + linkStyle + linktemplate + @"</head>",RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return inHTML;
        }
        public class StyleClass
        {
            private string _name = string.Empty;
            public string Name
            {
                get { return _name; }
                set { _name = value; }
            }

            private SortedList<string, string> _attributes = new SortedList<string, string>();
            public SortedList<string, string> Attributes
            {
                get { return _attributes; }
                set { _attributes = value; }
            }
        }
    }
}
