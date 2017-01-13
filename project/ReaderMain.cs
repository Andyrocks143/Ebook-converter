using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using wrd = Microsoft.Office.Interop.Word;
using System.IO;
namespace WindowsFormsApplication4
{
  
    public partial class ReaderMain : Form
    {
        public ReaderMain()
        {
            InitializeComponent();
        }

        private void ReaderMain_Load(object sender, EventArgs e)
        {
            
        }

        private void expPanel1_Load(object sender, EventArgs e)
        {

        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            ////
        }
        private string AddChaterBookmark(string inHTML, string id)
        {
            Regex oRegx= new Regex(@"\<body(.*?)\>(.*?)\<\/body\>",RegexOptions.IgnoreCase | RegexOptions.Singleline);
            MatchCollection oMats = oRegx.Matches(inHTML);
            if (oMats.Count > 0)
            {
                oRegx= new Regex(@"(\<p.*?\>)(.*?)(\<\/p\>)",RegexOptions.IgnoreCase | RegexOptions.Singleline);
                oMats = oRegx.Matches(oMats[0].Value);
                if(oMats.Count > 0)
                {
                    inHTML = oRegx.Replace(inHTML, "$1<a name='ch" + id + "'>$2</a>$3");
                }
            }
            return inHTML;
        }
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            if (txtBookID.Text == "" | txtAuthorname.Text == "" | txtPublisher.Text == "" | txtBookname.Text == "")
            {
                MessageBox.Show("Details missing");

            }
            else
            {
                string files = "";
                string isbn = txtBookID.Text;
                string stitle = txtBookname.Text;
                string sauthor = txtAuthorname.Text;
                string spublisher = txtPublisher.Text;
                if (expPanel1.FMFileList != null)
                    foreach (string s in expPanel1.FMFileList)
                        files += s + Environment.NewLine;
                if (expPanel1.BMFileList != null)
                    foreach (string s in expPanel1.BMFileList)
                        files += s + Environment.NewLine;
                if (expPanel1.EMFileList != null)
                    foreach (string s in expPanel1.EMFileList)
                        files += s + Environment.NewLine;
                string[] a = files.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                if (a.Length == 0)
                {
                    MessageBox.Show("Please add the files");
                    return;
                }
                string[] htmlFiles = new string[a.Length];
                string destPath = System.IO.Path.GetDirectoryName(a[0]) + "\\Epubconv\\" + isbn + "\\OEBPS\\HTML";
                string dest1Path = System.IO.Path.GetDirectoryName(a[0]) + "\\Epubconv\\" + isbn + "\\OEBPS\\HTML";
                string containerPath = System.IO.Path.GetDirectoryName(a[0]) + "\\Epubconv\\" + isbn + "\\META-INF";

                if (System.IO.Directory.Exists(destPath) == false)
                {
                    System.IO.Directory.CreateDirectory(destPath);
                }

                if (System.IO.Directory.Exists(containerPath) == false)
                {
                    System.IO.Directory.CreateDirectory(containerPath);
                }

                wrd.Application wrdApp = new wrd.Application();
                try
                {
                    wrdApp.Visible = false;
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = a.Length;
                    wrd.Document doc = wrdApp.Documents.Add();
                    string dest = "EpubCss";
                    object formt = wrd.WdSaveFormat.wdFormatFilteredHTML;
                    string final = System.IO.Path.Combine(destPath, dest + ".html");

                    //Merging the files to get the whole style sheet
                    for (int i = 0; i < a.Length; i++)
                    {
                        wrd.Range insertRng = doc.Content.Duplicate;
                        insertRng.InsertAfter(Convert.ToChar(13).ToString());
                        insertRng = doc.Content.Duplicate;
                        insertRng.Collapse(wrd.WdCollapseDirection.wdCollapseEnd);
                        insertRng.InsertFile(a[i]);
                        if (progressBar1.Value + 1 <= progressBar1.Maximum)
                        {
                            progressBar1.Value = progressBar1.Value + 1;
                        }
                        this.Refresh();
                    }
                    doc.SaveAs(final, formt);
                    doc.Close();
                    //Ends here

                    string htmlstring = System.IO.File.ReadAllText(final, Encoding.UTF8);
                    for (int i = 0; i < a.Length; i++)
                    {
                        label2.Text = "Opening" + System.IO.Path.GetFileNameWithoutExtension(a[i]);
                        doc = wrdApp.Documents.Open(a[i]);
                        wrdApp.Visible = false;
                        dest = System.IO.Path.GetFileNameWithoutExtension(a[i]);
                        formt = wrd.WdSaveFormat.wdFormatFilteredHTML;
                        final = System.IO.Path.Combine(destPath, dest + ".html");

                        AutoProcess.clsCleanup oCls = new AutoProcess.clsCleanup();
                        oCls.SetWordDoc = doc;
                        label2.Text = "Removing breaks in" + System.IO.Path.GetFileNameWithoutExtension(a[i]);
                        oCls.ToRemoveBreaks(true, true, true, true);
                        label2.Text = "Removing blank paragraphs in" + System.IO.Path.GetFileNameWithoutExtension(a[i]);
                        oCls.ToRemoveBlankParagraphs();
                        label2.Text = "Removing comments in" + System.IO.Path.GetFileNameWithoutExtension(a[i]);
                        oCls.ToRemoveComments();
                        oCls = null;
                        doc.SaveAs(final, formt);
                        htmlFiles[i] = final;
                        doc.Close();
                        if (progressBar1.Value + 1 <= progressBar1.Maximum)
                        {
                            progressBar1.Value = progressBar1.Value + 1;
                        }
                        this.Refresh();
                    }


                    string stylesheetname = "Stylesheet.css";

                    //Cleanup the final HTML files Using tidy
                    this.Refresh();
                    progressBar1.Maximum = htmlFiles.Length; progressBar1.Value = 0; progressBar1.Minimum = 0;
                    List<string> Chaplist = new List<string>();
                    if (htmlFiles.Length != 0)
                    {
                        for (int i = 0; i < htmlFiles.Length; i++)
                        {
                            if (System.IO.File.Exists(htmlFiles[i]))
                            {
                                label2.Text = "Cleaning using tidy";
                                Tidy.Document oDoc = new Tidy.Document();
                                System.IO.StreamReader oreader = new System.IO.StreamReader(htmlFiles[i], Encoding.Default);
                                string HTMLText = oreader.ReadToEnd();
                                oreader.Close();

                                oDoc.SetOptInt(Tidy.TidyOptionId.TidyHideComments, 1);
                                oDoc.SetOptInt(Tidy.TidyOptionId.TidyMakeClean, 1);
                                oDoc.SetOptInt(Tidy.TidyOptionId.TidyQuoteMarks, 0);
                                oDoc.SetOptInt(Tidy.TidyOptionId.TidyWord2000, 1);
                                int wCoding = oDoc.GetOptInt(Tidy.TidyOptionId.TidyCharEncoding);
                                oDoc.SetOptInt(Tidy.TidyOptionId.TidyCharEncoding, wCoding);

                                oDoc.ParseString(HTMLText);
                                oDoc.CleanAndRepair();
                                HTMLText = oDoc.SaveString();

                                CssParser oParser = new CssParser();
                                HTMLText = oParser.CleanHtml(HTMLText, htmlFiles[i]);

                                HTMLText = oParser.addStylesheetLinking(stylesheetname, stitle, HTMLText);

                                //Add bookmark
                                HTMLText = AddChaterBookmark(HTMLText, i.ToString());
                                Chaplist.Add(htmlFiles[i] + "#" + "ch" + i.ToString());
                                //Ends here


                                System.IO.StreamWriter writer = new System.IO.StreamWriter(htmlFiles[i], false, Encoding.Default);
                                writer.Write(HTMLText);
                                writer.Close();

                                progressBar1.Value = progressBar1.Value + 1;
                            }
                        }
                    }

                    //Creating CSS
                    {
                        progressBar1.Name = "Generating CSS StyleSheet";
                        Tidy.Document oDoc = new Tidy.Document();
                        oDoc.SetOptInt(Tidy.TidyOptionId.TidyHideComments, 1);
                        oDoc.SetOptInt(Tidy.TidyOptionId.TidyMakeClean, 1);
                        oDoc.SetOptInt(Tidy.TidyOptionId.TidyQuoteMarks, 0);
                        oDoc.SetOptInt(Tidy.TidyOptionId.TidyWord2000, 1);
                        int wCoding = oDoc.GetOptInt(Tidy.TidyOptionId.TidyCharEncoding);
                        oDoc.SetOptInt(Tidy.TidyOptionId.TidyCharEncoding, wCoding);
                        oDoc.ParseFile(destPath + "\\Epubcss.html");
                        oDoc.CleanAndRepair();
                        oDoc.SaveFile(destPath + "\\Epubcss.html");

                        string target = System.IO.Path.Combine(destPath + "\\Epubcss.html");                        
                        System.IO.StreamReader or = new System.IO.StreamReader(target, Encoding.Default);
                        string ht = or.ReadToEnd();
                        
                        ht = ht.Replace("/*", "");
                        ht = ht.Replace("*/", "");
                        Regex reg = new Regex(@"\<style(.*?)\>(.*?)<\/style\>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                        MatchCollection match = reg.Matches(ht);
                        if (match.Count > 0)
                        {
                            File.WriteAllText(destPath + "\\" + stylesheetname, match[match.Count - 1].Value);
                        }
                        or.Close();

                        System.IO.File.Delete(destPath + "\\Epubcss.html");

                        //Add TOC HTML
                        string tocHTMLname = "toc.html";
                        target = System.IO.Path.Combine(destPath + "\\" + tocHTMLname);
                        System.IO.StreamWriter writer = new System.IO.StreamWriter(target, false, Encoding.Default);
                        writer.WriteLine("<html><body>");
                        int id = 0;
                        foreach (string book in Chaplist)
                        {
                            id = id + 1;
                            string tmpTxt = book;
                            char[] oChar = new char[] { '#' };
                            string chapname = tmpTxt.Split(oChar)[0];
                            chapname = Path.GetFileName(chapname);
                            tmpTxt = tmpTxt.Split(oChar)[1];
                            writer.WriteLine("<a href='" + chapname + "#" + tmpTxt + "'>" + "Chapter " + id.ToString() + "</a>");
                        }
                        writer.WriteLine("</body></html>");
                        writer.Close();
                        //Ends here

                        if (File.Exists(Application.StartupPath + "\\" + "page-template.xpgt"))
                        {
                            File.Copy(Application.StartupPath + "\\" + "page-template.xpgt", destPath + "\\page-template.xpgt", true);
                        }


                        //OPF creation
                        string orgOPFPath = Application.StartupPath + "\\Sample.opf";
                        StreamReader oReader = new StreamReader(orgOPFPath);
                        string tmptxt = oReader.ReadToEnd();
                        oReader.Close();


                        Regex sRegex = new Regex(@"\<\<(.*?)\>\>(\>)?", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                        MatchCollection sMats = sRegex.Matches(tmptxt);
                        foreach (Match sMat in sMats)
                        {

                            //Please add

                            if (sMat.Value.ToLower().Contains("<<title>>"))
                            {
                                tmptxt = tmptxt.Replace(sMat.Value, stitle);
                            }
                            if (sMat.Value.ToLower().Contains("<<author>>"))
                            {
                                tmptxt = tmptxt.Replace(sMat.Value,sauthor);
                            }
                            if (sMat.Value.ToLower().Contains("<<publisher>>"))
                            {
                                tmptxt = tmptxt.Replace(sMat.Value,spublisher);
                            }



                            if (sMat.Value.ToLower().Contains("<<px>>"))
                            {
                                tmptxt = tmptxt.Replace(sMat.Value, isbn);
                            }
                            if (sMat.Value.ToLower().Contains("<<stylesheet>>"))
                            {
                                tmptxt = tmptxt.Replace(sMat.Value, stylesheetname);
                            }
                            if (sMat.Value.ToLower().Contains("listchapter"))
                            {
                                string oText = sMat.Value.ToLower();
                                string totalText = "";
                                foreach (string chap in Chaplist)
                                {
                                    char[] splChars = new char[] { '#' };
                                    string t = chap.Split(splChars)[0];
                                    string xid = chap.Split(splChars)[1];

                                    t = Path.GetFileName(t);
                                    t = oText.Replace("ListChapter".ToLower(), t);
                                    t = t.Replace("^^id^^", xid);
                                    t = t.Replace("<<", ""); t = t.Replace(">>", "");
                                    if (totalText == "")
                                        totalText = totalText + t;
                                    else
                                        totalText = totalText + Convert.ToChar(13) + t;
                                }
                                tmptxt = tmptxt.Replace(sMat.Value, totalText);
                            }
                            if (sMat.Value.ToLower().Contains("listspine"))
                            {
                                string oText = sMat.Value.ToLower();
                                string totalText = "";
                                foreach (string chap in Chaplist)
                                {
                                    char[] splChars = new char[] { '#' };
                                    string t = chap.Split(splChars)[1];
                                    t = oText.Replace("listspine", t);
                                    t = t.Replace("<<", ""); t = t.Replace(">>", "");
                                    if (totalText == "")
                                        totalText = totalText + t;
                                    else
                                        totalText = totalText + Convert.ToChar(13) + t;
                                }
                                tmptxt = tmptxt.Replace(sMat.Value, totalText);
                            }
                            if (sMat.Value.ToLower().Contains("<<chapterone>>"))
                            {
                                char[] splChars = new char[] { '#' };
                                string t = Chaplist[0];
                                t = t.Split(splChars)[0];
                                t = Path.GetFileName(t) + "#" + Chaplist[0].Split(splChars)[1];
                                tmptxt = tmptxt.Replace(sMat.Value, t);
                            }
                            if (sMat.Value.ToLower().Contains("<<toc>>"))
                            {
                                tmptxt = tmptxt.Replace(sMat.Value, tocHTMLname);
                            }
                        }
                        StreamWriter oWriter = new StreamWriter(Directory.GetParent(destPath).ToString() + "\\" + isbn + ".opf");
                        oWriter.Write(tmptxt);
                        oWriter.Close();
                        //Ends here

                        //containerPath
                        if (File.Exists(Application.StartupPath + "\\container.xml"))
                        {
                            StreamReader oRd = new StreamReader(Application.StartupPath + "\\container.xml", Encoding.Default);
                            string conText = oRd.ReadToEnd();
                            oRd.Close();
                            conText = conText.Replace("$opfname$", isbn + ".opf");
                            StreamWriter owriter = new StreamWriter(containerPath + "\\container.xml");
                            owriter.Write(conText);
                            owriter.Close();
                        }

                        oWriter = new StreamWriter(Directory.GetParent(containerPath).ToString() + "\\mimetype");
                        oWriter.Write("application/epub+zip");
                        oWriter.Close();


                        progressBar1.Name = "Generating toc";
                        String orgNCXPath = Application.StartupPath + "\\Sample.ncx";
                        StreamReader oRr = new StreamReader(orgNCXPath);
                        string Readtxt = oRr.ReadToEnd();
                        oRr.Close();


                        Readtxt = Readtxt.Replace("^Title^", stitle);
                        Readtxt = Readtxt.Replace("^Author^", sauthor);

                        int i = 1;
                        Regex nRegex = new Regex(@"\<navPoint(.*?)\>(.*?)\<\/navPoint\>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                        MatchCollection nmatch = nRegex.Matches(Readtxt);
                        if (nmatch.Count > 0)
                        {
                            string txt = "";
                            string orgMatch = nmatch[0].Value;
                            string totalTxt = "";
                            foreach (string chap in Chaplist)
                            {
                                char[] splChars = new char[] { '#' };
                                string t = chap.Split(splChars)[0];
                                txt = orgMatch.Replace("^num^", i.ToString());
                                txt = txt.Replace("^chapters^", Path.GetFileName(t) + "#" + chap.Split(splChars)[1]);
                                totalTxt = totalTxt + @"<navPoint playorder='" + i + @"'>" + txt + @"</navPoint>";
                                i++;
                            }

                            Readtxt = Readtxt.Replace(orgMatch, totalTxt);
                        }

                        StreamWriter oW = new StreamWriter(Directory.GetParent(destPath).ToString() + "\\" + "toc.ncx");
                        oW.Write(Readtxt);
                        oW.Close();

                        label2.Text = "Completed";
                        MessageBox.Show("Completed");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    wrdApp.Quit();
                }
            }
        }
       
    }
}
