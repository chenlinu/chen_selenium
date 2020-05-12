using HtmlAgilityPack;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SeleniumSample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        const string _BAIDU_URL = "https://www.baidu.com";
        SeleniumHelper _webA = null;
        List<SeleniumHelper> _webLst = new List<SeleniumHelper>();//用于存放多个浏览器的集合

        private void AppendText(string txt)
        {
            richTextBox1.AppendText(string.Format(" {1} {0}\r\n", txt, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.SelectionLength = 0;
            richTextBox1.Focus();
        }

        private void CreateOneWeb()
        {
            if (_webA == null)
                _webA = new SeleniumHelper(Browsers.Chrome);//Browsers.Chrome 谷歌浏览器 Browsers.Firefox 火狐浏览器
        }

        private void 创建一个浏览器NToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();
            _webA.GoToUrl(_BAIDU_URL);
            AppendText("创建一个浏览器======完成！");
        }

        private void 创建多个浏览器MToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (var item in _webLst)//如果集合中有已经创建的浏览器，先将其关闭
            {
                item.Quit();
            }
            for (int i = 0; i < 5; i++)
            {
                _webLst.Add(new SeleniumHelper(Browsers.Chrome));
            }

            AppendText("创建多个浏览器======完成！");
        }
        SeleniumHelperCookie _webC = null, _webD = null;
        private void 创建一个能保存cookie信息的浏览器ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(Application.StartupPath + "//cacheC") == false)
                Directory.CreateDirectory(Application.StartupPath + "//cacheC");//创建一个目录，用于存储浏览器文件

            _webC = new SeleniumHelperCookie(Browsers.Chrome, Application.StartupPath + "//cacheC");// cacheA 是指定浏览器文件存放位置
            _webC.GoToUrl("https://weibo.com");

            AppendText("创建或打开一个能保存cookie信息的浏览器======完成！你可以登录一个账号，退出程序再通过当前菜单打开浏览器，看看账号是否已登录。");
        }

        private void 创建三个能保存cookie信息的浏览器NToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(Application.StartupPath + "//cacheC") == false)
                Directory.CreateDirectory(Application.StartupPath + "//cacheC");//创建一个目录，用于存储浏览器文件

            _webC = new SeleniumHelperCookie(Browsers.Chrome, Application.StartupPath + "//cacheC");// cacheA 是指定浏览器文件存放位置
            _webC.GoToUrl("https://weibo.com");

            if (Directory.Exists(Application.StartupPath + "//cacheD") == false)
                Directory.CreateDirectory(Application.StartupPath + "//cacheD");//创建一个目录，用于存储浏览器文件

            _webD = new SeleniumHelperCookie(Browsers.Chrome, Application.StartupPath + "//cacheD");// cacheA 是指定浏览器文件存放位置
            _webD.GoToUrl(_BAIDU_URL);

            AppendText("创建或打开多个能保存cookie信息的浏览器======完成！");
        }

        private void alter消息框MToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl("http://2dzcx.cdpf.org.cn/cdpf");
            _webA.FindElementById("name").SendKeys("王天天");
            _webA.FindElementById("idcard").SendKeys("888888888888888888");
            _webA.FindElementById("user_yzm_image").SendKeys("6666");
            _webA.FindElementById("login_sub").Click();

            try
            {
                Commons.Wait(3);
                IAlert confirm = _webA.wd.SwitchTo().Alert();
                AppendText("收到消息：" + confirm.Text);
                Commons.Wait(3);
                confirm.Accept();
            }
            catch (Exception)
            {
            }
            AppendText("alter消息框======完成！");
        }

        private void iframe切换IToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl("https://music.163.com/#/discover/djradio");

            IWebElement frame = _webA.FindElementById("g_iframe");
            _webA.wd.SwitchTo().Frame(frame);//切换到Iframe
                                             //_webA.wd.SwitchTo().Frame("g_iframe");//也可以
            AppendText("iframe切换======完成！");
            AppendText("即将在iframe页面中点击第二页");
            Commons.Wait(3);
            if (_webA.FindElementByXPath("//span[@data-index='2']") != null)
                _webA.FindElementByXPath("//span[@data-index='2']").Click();
            AppendText("iframe切换操作======完成！");
        }

        private void 选项卡切换TToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl("http://news.baidu.com/");

            //打开某一篇新闻
            _webA.FindElementsByXPathName("//a[@class='a3']")[0].Click();
            AppendText("即将切换到新打开的标签页！");
            Commons.Wait(2);

            String winHandleBefore = _webA.wd.CurrentWindowHandle;//先获取当前页面的Handle
            foreach (var item in _webA.wd.WindowHandles)
            {
                if (winHandleBefore != item)
                {
                    _webA.wd.SwitchTo().Window(item);//再转入到新的标签页
                                                     //TODO 可对新标签页进行操作
                    Commons.Wait(2);
                    string txt = _webA.FindElementByXPath("//body").Text.Replace(" ", "").Replace("\r\n", "").Replace("\t", "");
                    AppendText("获取到新标签页的内容有：" + txt);
                }
            }

            if (_webA.wd.WindowHandles.Count >= 2)
            {
                AppendText("即将关闭新标签页，并回到主页！");
                Commons.Wait(2);
                _webA.wd.Close();
                _webA.wd.SwitchTo().Window(winHandleBefore);

                AppendText("标签页切换操作======完成！");
            }
        }

        private void 填写内容WToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl(_BAIDU_URL);

            Commons.Wait(2);
            _webA.FindElementById("kw").SendKeys("Hi Selenium");
            AppendText("填写内容======完成！");
        }

        private void 点击元素CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl(_BAIDU_URL);

            Commons.Wait(2);
            _webA.FindElementByXPath("//a[contains(text(),'新闻')]").Click();
            //点击一个A标签中包含“新闻”的元素
            AppendText("点击元素======完成！");
        }

        private void 注入或执行JSIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl(_BAIDU_URL);

            string js = "document.getElementById('kw').style = 'border:6px solid red;';";
            ((IJavaScriptExecutor)_webA.wd).ExecuteScript(js);

            js = "document.getElementById('su').value = 'Button X';";
            ((IJavaScriptExecutor)_webA.wd).ExecuteScript(js);
            AppendText("注入或执行JS======完成！");
        }

        private void 识别验证码VToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl("http://2dzcx.cdpf.org.cn/cdpf");

            if (System.IO.Directory.Exists(Application.StartupPath + "\\tmp\\") == false)
                System.IO.Directory.CreateDirectory(Application.StartupPath + "\\tmp\\");

            string img_path = Application.StartupPath + "\\tmp\\" + DateTime.Now.ToFileTime() + ".jpg";
            IWebElement ele = _webA.FindElementByXPath("//img[@alt='中国残疾人联合会']");//获取验证码图片对象
            int x = ele.Location.X, y = ele.Location.Y, width = ele.Size.Width, height = ele.Size.Height;

            _webA.TakeScreenshot(img_path);
            Bitmap source = new Bitmap(img_path);
            Bitmap CroppedImage = source.Clone(new System.Drawing.Rectangle(x, y, width, height), source.PixelFormat);
            CroppedImage.Save(img_path + ".jpg");//保存验证码图片

            //该示例所使用的是调用“联众打码”的接口，你也可以通过自己的方式去识别。
            string result = FastVerCode.VerCode.RecYZM_A_2(img_path + ".jpg", 1108, 2, 9, "你的账号", "你的密码", "你的Token");
            if (result.Contains("TimeOut"))
                return;
            string val = result.Split('|')[0];
            _webA.FindElementById("user_yzm_image").SendKeys(val);

            AppendText("识别验证码======完成！");
        }

        private void 通过浏览器进程ID关闭浏览器CToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SeleniumHelper.KillProcessAndChildren(_webA.ProcessID);
            AppendText("通过进程关闭浏览器======完成！");
        }

        private void 获取单个元素LToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl(_BAIDU_URL);

            IWebElement ele = _webA.FindElementByXPath("//span[@class='hot-refresh-text']");
            //ele = _webA.FindElementByXPath("//*[@id='hotsearch-refresh-btn']/span"); 这样也行
            AppendText("获取到Span标签：" + ele.Text);
            AppendText("获取单个元素======完成！");
        }

        private void 获取多个元素LToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl("http://news.baidu.com/");

            IList<IWebElement> lst = _webA.FindElementsByXPathName("//a[@class='a3']");
            foreach (var item in lst)
            {
                AppendText(item.Text);
            }
            AppendText("获取多个元素======完成！");
        }

        private void htmlDocumentDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl(_BAIDU_URL);

            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(_webA.GetSource());
            AppendText("htmlDocument加载的内容：" + document.DocumentNode.InnerText.Replace("\r\n", "").Replace("\t", "").Replace(" ", ""));

            AppendText("htmlDocument加载页面======完成！");
        }

        private void htmlNodeNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl(_BAIDU_URL);

            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(_webA.GetSource());
            HtmlNode htmlNode = document.DocumentNode.SelectSingleNode("//input[@id='su']");
            AppendText("htmlNode获取到按钮：" + htmlNode.InnerText);
            AppendText("htmlNode获取到按钮的XPath：" + htmlNode.XPath);

            AppendText("htmlNode获取对象======完成！");
        }

        private void htmlNodeCollectionCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl(_BAIDU_URL);

            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
            document.LoadHtml(_webA.GetSource());
            HtmlNodeCollection htmlNodes = document.DocumentNode.SelectNodes("//a");

            AppendText("HtmlNodeCollection获取到的集合：");
            foreach (var item in htmlNodes)
            {
                AppendText("文本：" + item.InnerText + " & XPath：" + item.XPath);
            }

            AppendText("HtmlNodeCollection获取集合======完成！");
        }

        private void 退出XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_webA != null)
                _webA.Quit();

            Process.GetCurrentProcess().Kill();
        }

        private void CloseWeb()
        {
            if (_webA != null)
                _webA.Quit();
            if (_webC != null)
                _webC.Quit();
            if (_webD != null)
                _webD.Quit();
            foreach (var item in _webLst)
            {
                item.Quit();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            CloseWeb();
        }

        private void 打开一个新的标签页OToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateOneWeb();//检测浏览器是否已经实例化好了
            _webA.GoToUrl(_BAIDU_URL);

            AppendText("即将打开一个新的标签页");
            Commons.Wait(2);
            _webA.ExecuteJS("window.open('about: blank','_blank');");
            String winHandleBefore = _webA.wd.CurrentWindowHandle;//先获取当前页面的Handle
            foreach (var item in _webA.wd.WindowHandles)
            {
                if (winHandleBefore != item)
                {
                    _webA.wd.SwitchTo().Window(item);//再转入到新的标签页
                    _webA.GoToUrl("https://www.sogou.com/");
                }
            }
            AppendText("打开一个新的标签页======完成！");
        }
    }
}
