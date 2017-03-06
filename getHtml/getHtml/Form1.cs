using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using StanSoft;
using word;
using System.Text.RegularExpressions;

namespace getHtml
{
    public partial class Form1 : Form
    {
        //初始化
        public Form1()
        {
            InitializeComponent();
            this.backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            this.backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            this.backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
        }
        //通过响应消息，来处理界面的显示工作
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Value = e.ProgressPercentage;
            this.label2.Text = e.UserState.ToString();
            this.label2.Update();
        }

        //这里是后台工作完成后的消息处理，可以在这里进行后续的处理工作。
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("生成成功！");
            this.progressBar1.Visible = false;
            this.label2.Text = "";
        }

        //这里，就是后台进程开始工作时，调用工作函数的地方。你可以把你现有的处理函数写在这儿。
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            work(this.backgroundWorker1, e);
        }
        /// <summary>
        /// 处理函数
        /// </summary>
        /// <param name="bk"></param>
        /// <returns></returns>
        private bool work(BackgroundWorker bk, DoWorkEventArgs e)
        {
            string html = e.Argument.ToString();
            try
            {

                bk.ReportProgress(20, "20%");



                bk.ReportProgress(30, "30%");

                // article对象包含Title(标题)，PublishDate(发布日期)，Content(正文)和ContentWithTags(带标签正文)四个属性
                Article article = Html2Article.GetArticle(html);

                bk.ReportProgress(50, "50%");

                //1,首先需要载入模板
                Report report = new Report();
                report.CreateNewDocument(Application.StartupPath + "\\demo.doc"); //模板路径

                bk.ReportProgress(70, "70%");

                //2,插入一个值
                report.InsertValue("title", article.Title);
                report.InsertValue("date", article.PublishDate.ToString());
                report.InsertValue("content", article.Content);

                bk.ReportProgress(80, "80%");

                report.SaveDocument(Application.StartupPath +
                    "\\..\\..\\..\\document" + "\\爬下的东西" +
                    DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") + ".doc"); //文档路径
                bk.ReportProgress(100, "100%");

                //MessageBox.Show("生成成功");
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            //处理的过程中，通过这个函数，向主线程报告处理进度，最好是折算成百分比，与外边的进度条的最大值必须要对应。这里，我没有折算，而是把界面线程的进度条最大值调整为与这里的总数一致。


        }
        /// <summary>
        /// getHtml按钮事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void getHtml_Click(object sender, EventArgs e)
        {
            try
            {
                string URL = getURL.Text;
                getURL.Text = "";
                string html = HTMLHelper.GetHTMLByUrl(URL, "utf-8");
                string RegexStr = string.Empty;
                RegexStr = "charset=[\\S]+\">";          //要匹配的字符串
                Match mt = Regex.Match(html, RegexStr);
                if (mt.Value != "")
                {
                    switch (mt.Value)
                    {
                        case "charset=utf-8":
                            break;
                        case "charset=utf-8\"":
                            break;
                        case "charset=\"utf-8\"":
                            break;
                        case "charset=\"utf-8\">":
                            break;
                        case "charset=\"utf-8\"\\>":
                            break;

                        case "charset=gb2312":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb2312");
                            break;
                        case "charset=gb2312\"":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb2312");
                            break;
                        case "charset=\"gb2312\"":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb2312");
                            break;
                        case "charset=\"gb2312\">":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb2312");
                            break;
                        case "charset=\"gb2312\"\\>":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb2312");
                            break;

                        case "charset=gb18030":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                        case "charset=gb18030\"":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                        case "charset=\"gb18030\"":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                        case "charset=\"gb18030\">":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                        case "charset=\"gb18030\"\\>":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;

                        case "charset=gbk":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                        case "charset=gbk\"":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                        case "charset=\"gbk\"":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                        case "charset=\"gbk\">":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                        case "charset=\"gbk\"\\>":
                            html = HTMLHelper.GetHTMLByUrl(URL, "gb18030");
                            break;
                    }
                }
                if (html == null)
                    MessageBox.Show("url错误！");
                else
                {
                    this.progressBar1.Visible = true;
                    this.backgroundWorker1.RunWorkerAsync(html);
                }

            }
            catch (Exception)
            {

                MessageBox.Show("url错误！");
            }

        }
    }
}
