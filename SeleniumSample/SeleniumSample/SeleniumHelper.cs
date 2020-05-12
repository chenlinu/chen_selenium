
using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Safari;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QA = OpenQA.Selenium;
using UI = OpenQA.Selenium.Support.UI;

namespace SeleniumSample
{
    public class SeleniumHelper
    {
         public QA.IWebDriver wd = null;
        public Browsers browser = Browsers.Chrome;
        public QA.DriverService ds = null;
        bool _IsLoadPicture = true; bool _IsLoadJS = true; bool _IsUseNewProxy = false;
        public int ProcessID = 0;

        public SeleniumHelper(Browsers theBrowser)
        {
            this.browser = theBrowser;
            wd = InitWebDriver();
        }
        int _timeout = 60;
        public SeleniumHelper(Browsers theBrowser, int timeout)
        {
            this.browser = theBrowser;
            _timeout = timeout;
            wd = InitWebDriver();
        }
        public SeleniumHelper(Browsers theBrowser, bool IsUseNewProxy)
        {
            this.browser = theBrowser;
            _IsUseNewProxy = IsUseNewProxy;
            wd = InitWebDriver();
        }
        string _doproxy = "1";
        public SeleniumHelper(Browsers theBrowser, string doproxy)
        {
            this.browser = theBrowser;
            _doproxy = doproxy;
            wd = InitWebDriver();
        }
        public SeleniumHelper(Browsers theBrowser, bool IsLoadPicture, bool IsLoadJS)
        {
            this.browser = theBrowser;
            _IsLoadPicture = IsLoadPicture;
            _IsLoadJS = IsLoadJS;
            wd = InitWebDriver();
        }
        private QA.IWebDriver InitWebDriver()
        {
            QA.IWebDriver theDriver = null;
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        InternetExplorerDriverService driverService = InternetExplorerDriverService.CreateDefaultService(Application.StartupPath + "\\drivers\\");
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;
                        QA.IE.InternetExplorerOptions _ieOptions = new QA.IE.InternetExplorerOptions();
                        _ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                        theDriver = new QA.IE.InternetExplorerDriver(driverService, _ieOptions);
                        ProcessID = driverService.ProcessId;
                    }; break;
                case Browsers.PhantomJS:
                    {
                        PhantomJSDriverService driverService = PhantomJSDriverService.CreateDefaultService(Application.StartupPath + "\\Drivers\\");
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;

                        theDriver = new QA.PhantomJS.PhantomJSDriver(driverService);
                    }; break;
                case Browsers.Chrome:
                    {
                        ChromeDriverService driverService = ChromeDriverService.CreateDefaultService(Application.StartupPath + "\\drivers\\");
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;

                        ChromeOptions options = new QA.Chrome.ChromeOptions();
                        options.AddUserProfilePreference("profile.managed_default_content_settings.images", _IsLoadPicture ? 1 : 2);
                        options.AddUserProfilePreference("profile.managed_default_content_settings.javascript", _IsLoadJS ? 1 : 2);

                        //options.AddArgument(@"--user-data-dir=" + cache_dir);
                        //string dir = string.Format(@"user-data-dir={0}", ConfigManager.GetInstance().UserDataDir);
                        //options.AddArguments(dir);

                        //options.AddArgument("--no-sandbox");
                        //options.AddArgument("--disable-dev-shm-usage");
                        //options.AddArguments("--disable-extensions"); // to disable extension
                        //options.AddArguments("--disable-notifications"); // to disable notification
                        //options.AddArguments("--disable-application-cache"); // to disable cache
                        try
                        {
                            if (_timeout == 60)
                                theDriver = new QA.Chrome.ChromeDriver(driverService, options, new TimeSpan(0, 0, 40));
                            else
                                theDriver = new QA.Chrome.ChromeDriver(driverService, options, new TimeSpan(0, 0, _timeout));
                        }
                        catch (Exception ex)
                        {
                        }
                        ProcessID = driverService.ProcessId;
                    }; break;
                case Browsers.Firefox:
                    {
                        var driverService = FirefoxDriverService.CreateDefaultService(Application.StartupPath + "\\drivers\\");
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;

                        FirefoxProfile profile = new FirefoxProfile();
                        try
                        {
                            if (_doproxy == "1")
                            {
                                string proxy = "";
                                try
                                {
                                    if (_IsUseNewProxy == false)
                                    {
                                        proxy = GetProxyA();
                                    }
                                    else
                                    {
                                        //TO DO 获取芝麻代理
                                        hi.URL = "http:......?你的代理地址";// ConfigManager.GetInstance().ProxyUrl;
                                        hr = hh.GetContent(hi);
                                        if (hr.StatusCode == System.Net.HttpStatusCode.OK)
                                        {
                                            if (hr.Content.Contains("您的套餐余量为0"))
                                                proxy = "";
                                            if (hr.Content.Contains("success") == false)
                                                proxy = "";

                                            JObject j = JObject.Parse(hr.Content);
                                            foreach (var item in j)
                                            {
                                                foreach (var itemA in item.Value)
                                                {
                                                    if (itemA.ToString().Contains("expire_time"))
                                                    {
                                                        if (DateTime.Now.AddHours(2) < DateTime.Parse(itemA["expire_time"].ToString()))
                                                        {
                                                            proxy = itemA["ip"].ToString() + ":" + itemA["port"].ToString();
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (proxy != "")
                                                    break;
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                }

                                if (proxy != "" && proxy.Contains(":"))
                                {
                                    OpenQA.Selenium.Proxy proxyF = new OpenQA.Selenium.Proxy();
                                    proxyF.HttpProxy = proxy;
                                    proxyF.FtpProxy = proxy;
                                    proxyF.SslProxy = proxy;
                                    profile.SetProxyPreferences(proxyF);
                                    // 使用代理
                                    profile.SetPreference("network.proxy.type", 1);
                                    //ProxyUser-通行证书 ProxyPass-通行密钥
                                    profile.SetPreference("username", "你的账号");
                                    profile.SetPreference("password", "你的密码");

                                    // 所有协议公用一种代理配置，如果单独配置，这项设置为false
                                    profile.SetPreference("network.proxy.share_proxy_settings", true);

                                    // 对于localhost的不用代理，这里必须要配置，否则无法和webdriver通讯
                                    profile.SetPreference("network.proxy.no_proxies_on", "localhost");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                        }

                        //profile.SetPreference("permissions.default.image", 2);
                        // 关掉flash  
                        profile.SetPreference("dom.ipc.plugins.enabled.libflashplayer.so", false);
                        FirefoxOptions options = new FirefoxOptions();
                        options.Profile = profile;
                        theDriver = new QA.Firefox.FirefoxDriver(driverService, options, TimeSpan.FromMinutes(1));
                        ProcessID = driverService.ProcessId;
                    }; break;
                case Browsers.Safari:
                    {
                        SafariDriverService driverService = SafariDriverService.CreateDefaultService(Application.StartupPath + "\\Drivers\\");
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;
                        theDriver = new QA.Safari.SafariDriver(driverService);
                        ProcessID = driverService.ProcessId;
                    }; break;
                default:
                    {
                        QA.IE.InternetExplorerOptions _ieOptions = new QA.IE.InternetExplorerOptions();
                        _ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                        theDriver = new QA.IE.InternetExplorerDriver(_ieOptions);

                    }; break;
            }
            return theDriver;
        }

        HttpHelper hh = new HttpHelper();
        HttpItem hi = new HttpItem();
        HttpResult hr = null;
        public string GetProxyA()
        {
            for (int k = 0; k < 6; k++)
            {
                try
                {
                  
                    hi.URL = "";// ConfigManager.GetInstance().ProxyUrlDLY;
                    HttpResult hr = hh.GetContent(hi);
                    if (hr.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        string[] arr = hr.Content.Split('\r');
                        if (arr.Length < 5)
                            continue;
                        return arr[0].Replace("\n", "");
                    }
                }
                catch (Exception ex)
                {
                }
            }
            return "";
        }
        #region public members
        /// <summary>
        /// Effects throughout the life of web driver
        /// Set once only if necessary
        /// </summary>
        /// <param name="seconds"></param>
        public void ImplicitlyWait(double seconds)
        {
            wd.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(seconds);
        }

        /// <summary>
        /// Wait for the expected condition is satisfied, return immediately
        /// </summary>
        /// <param name="expectedCondition"></param>
        public void WaitForPage(string title)
        {
            //UI.WebDriverWait _wait = new UI.WebDriverWait(wd, TimeSpan.FromSeconds(10));
            //_wait.Until((d) => { return d.Title.ToLower().StartsWith(title.ToLower()); });
            //to do
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="we"></param>
        public void WaitForElement(string id)
        {
            //UI.WebDriverWait _wait = new UI.WebDriverWait(wd, TimeSpan.FromSeconds(10));
            //_wait.Until((d) => { return OpenQA.Selenium.Support.UI.ExpectedConditions.ElementExists(QA.By.Id(id)); });
        }


        /// <summary>
        /// Load a new web page in current browser
        /// </summary>
        /// <param name="url"></param>
        public void GoToUrl(string url)
        {
            try
            {
                wd.Navigate().GoToUrl(url);
            }
            catch (Exception)
            {
                Commons.Wait(1);
            }
        }

        public void Refresh()
        {
            for (int i = 0; i < 6; i++)
            {
                try
                {
                    wd.Navigate().Refresh();
                    break;
                }
                catch (Exception)
                {
                    Commons.Wait(1);
                }
            }
        }

        public void Back()
        {
            try
            {
                wd.Navigate().Back();
            }
            catch (Exception)
            {
            }
        }

        public void Forward()
        {
            wd.Navigate().Forward();
        }

        /// <summary>
        /// Get the url of current browser window
        /// </summary>
        /// <returns></returns>
        public string GetUrl()
        {
            for (int i = 0; i < 14; i++)
            {
                try
                {
                    string url = wd.Url;
                    return url;
                }
                catch (Exception)
                {
                    Commons.Wait(1);
                }
            }
            return wd.Url;
        }

        /// <summary>
        /// Get page title of current browser window
        /// </summary>
        /// <returns></returns>
        public string GetPageTitle()
        {
            for (int i = 0; i < 14; i++)
            {
                try
                {
                    string Title = wd.Title;
                    return Title;
                }
                catch (Exception)
                {
                    Commons.Wait(1);
                }
            }
            return wd.Title;
        }

        /// <summary>
        /// Get all cookies defined in the current page
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, string> GetAllCookies()
        {
            Dictionary<string, string> cookies = new Dictionary<string, string>();
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        var allCookies = ((QA.IE.InternetExplorerDriver)wd).Manage().Cookies.AllCookies;
                        foreach (QA.Cookie cookie in allCookies)
                        {
                            cookies[cookie.Name] = cookie.Value;
                        }
                    }; break;
                case Browsers.Chrome:
                    {
                        var allCookies = ((QA.Chrome.ChromeDriver)wd).Manage().Cookies.AllCookies;
                        foreach (QA.Cookie cookie in allCookies)
                        {
                            cookies[cookie.Name] = cookie.Value;
                        }
                    }; break;
                case Browsers.Firefox:
                    {
                        var allCookies = ((QA.Firefox.FirefoxDriver)wd).Manage().Cookies.AllCookies;
                        foreach (QA.Cookie cookie in allCookies)
                        {
                            cookies[cookie.Name] = cookie.Value;
                        }
                    }; break;
                case Browsers.Safari:
                    {
                        var allCookies = ((QA.Safari.SafariDriver)wd).Manage().Cookies.AllCookies;
                        foreach (QA.Cookie cookie in allCookies)
                        {
                            cookies[cookie.Name] = cookie.Value;
                        }
                    }; break;
                default:
                    {
                        var allCookies = ((QA.IE.InternetExplorerDriver)wd).Manage().Cookies.AllCookies;
                        foreach (QA.Cookie cookie in allCookies)
                        {
                            cookies[cookie.Name] = cookie.Value;
                        }
                    }; break;
            }

            return cookies;
        }

        /// <summary>
        /// Delete all cookies from the page
        /// </summary>
        public void DeleteAllCookies()
        {
            try
            {
                switch (this.browser)
                {
                    case Browsers.IE:
                        {
                            ((QA.IE.InternetExplorerDriver)wd).Manage().Cookies.DeleteAllCookies();
                        }; break;
                    case Browsers.Chrome:
                        {
                            ((QA.Chrome.ChromeDriver)wd).Manage().Cookies.DeleteAllCookies();
                        }; break;
                    case Browsers.Firefox:
                        {
                            ((QA.Firefox.FirefoxDriver)wd).Manage().Cookies.DeleteAllCookies();
                        }; break;
                    case Browsers.Safari:
                        {
                            ((QA.Safari.SafariDriver)wd).Manage().Cookies.DeleteAllCookies();
                        }; break;
                    default:
                        {
                            ((QA.IE.InternetExplorerDriver)wd).Manage().Cookies.DeleteAllCookies();
                        }; break;
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Set focus to a browser window with a specified title
        /// </summary>
        /// <param name="title"></param>
        /// <param name="exactMatch"></param>
        public void GoToWindow(string title, bool exactMatch)
        {
            string theCurrent = wd.CurrentWindowHandle;
            IList<string> windows = wd.WindowHandles;
            if (exactMatch)
            {
                foreach (var window in windows)
                {
                    wd.SwitchTo().Window(window);
                    if (wd.Title.ToLower() == title.ToLower())
                    {
                        return;
                    }
                }
            }
            else
            {
                foreach (var window in windows)
                {
                    wd.SwitchTo().Window(window);
                    if (wd.Title.ToLower().Contains(title.ToLower()))
                    {
                        return;
                    }
                }
            }

            wd.SwitchTo().Window(theCurrent);
        }

        /// <summary>
        /// Set focus to a frame with a specified name
        /// </summary>
        /// <param name="name"></param>
        public void GoToFrame(string name)
        {
            QA.IWebElement theFrame = null;
            var frames = wd.FindElements(QA.By.TagName("iframe"));
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name").ToLower() == name.ToLower())
                {
                    theFrame = (QA.IWebElement)frame;
                    break;
                }
            }
            if (theFrame != null)
            {
                wd.SwitchTo().Frame(theFrame);
            }
        }

        public void GoToFrame(QA.IWebElement frame)
        {
            wd.SwitchTo().Frame(frame);
        }

        /// <summary>
        /// Switch to default after going to a frame
        /// </summary>
        public void GoToDefault()
        {
            wd.SwitchTo().DefaultContent();
        }

        /// <summary>
        /// Get the alert text
        /// </summary>
        /// <returns></returns>
        public string GetAlertString()
        {
            string theString = string.Empty;
            QA.IAlert alert = null;
            alert = wd.SwitchTo().Alert();
            if (alert != null)
            {
                theString = alert.Text;
            }
            return theString;
        }

        /// <summary>
        /// Accepts the alert
        /// </summary>
        public void AlertAccept()
        {
            QA.IAlert alert = null;
            alert = wd.SwitchTo().Alert();
            if (alert != null)
            {
                alert.Accept();
            }
        }

        /// <summary>
        /// Dismisses the alert
        /// </summary>
        public void AlertDismiss()
        {
            QA.IAlert alert = null;
            alert = wd.SwitchTo().Alert();
            if (alert != null)
            {
                alert.Dismiss();
            }
        }

        /// <summary>
        /// Move vertical scroll bar to bottom for the page
        /// </summary>
        public void PageScrollToBottom()
        {
            try
            {
                var js = "document.documentElement.scrollTop=990000";
                switch (this.browser)
                {
                    case Browsers.IE:
                        {
                            ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                        }; break;
                    case Browsers.Chrome:
                        {
                            ((QA.Chrome.ChromeDriver)wd).ExecuteScript(js, null);
                        }; break;
                    case Browsers.Firefox:
                        {
                            ((QA.Firefox.FirefoxDriver)wd).ExecuteScript(js, null);
                        }; break;
                    case Browsers.PhantomJS:
                        {
                            ((QA.PhantomJS.PhantomJSDriver)wd).ExecuteScript(js, null);
                        }; break;
                    case Browsers.Safari:
                        {
                            ((QA.Safari.SafariDriver)wd).ExecuteScript(js, null);
                        }; break;
                    default:
                        {
                            ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                        }; break;
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Move vertical scroll bar to bottom for the page
        /// </summary>
        public void PageScrollToTop()
        {
            var js = "document.documentElement.scrollTop=100";
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Chrome:
                    {
                        ((QA.Chrome.ChromeDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Firefox:
                    {
                        ((QA.Firefox.FirefoxDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.PhantomJS:
                    {
                        ((QA.PhantomJS.PhantomJSDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Safari:
                    {
                        ((QA.Safari.SafariDriver)wd).ExecuteScript(js, null);
                    }; break;
                default:
                    {
                        ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                    }; break;
            }
        }
        public Image GetEntireScreenshot()
        {
            ((QA.IJavaScriptExecutor)this.wd).ExecuteScript("window.scrollTo(0, 0)");
            // Get the total size of the page
            var totalWidth = (int)(long)((QA.IJavaScriptExecutor)this.wd).ExecuteScript("return document.body.offsetWidth"); //documentElement.scrollWidth");
            var totalHeight = (int)(long)((QA.IJavaScriptExecutor)wd).ExecuteScript("return  document.body.parentNode.scrollHeight");
            // Get the size of the viewport
            var viewportWidth = (int)(long)((QA.IJavaScriptExecutor)wd).ExecuteScript("return document.body.clientWidth"); //documentElement.scrollWidth");
            var viewportHeight = (int)(long)((QA.IJavaScriptExecutor)wd).ExecuteScript("return window.innerHeight"); //documentElement.scrollWidth");

            // We only care about taking multiple images together if it doesn't already fit
            if (totalWidth <= viewportWidth && totalHeight <= viewportHeight)
            {
                var screenshot = ((QA.ITakesScreenshot)wd).GetScreenshot();
                return ScreenshotToImage(screenshot);
            }
            // Split the screen in multiple Rectangles
            var rectangles = new List<Rectangle>();
            // Loop until the totalHeight is reached
            for (var y = 0; y < totalHeight; y += viewportHeight)
            {
                var newHeight = viewportHeight;
                // Fix if the height of the element is too big
                if (y + viewportHeight > totalHeight)
                {
                    newHeight = totalHeight - y;
                }
                // Loop until the totalWidth is reached
                for (var x = 0; x < totalWidth; x += viewportWidth)
                {
                    var newWidth = viewportWidth;
                    // Fix if the Width of the Element is too big
                    if (x + viewportWidth > totalWidth)
                    {
                        newWidth = totalWidth - x;
                    }
                    // Create and add the Rectangle
                    var currRect = new Rectangle(x, y, newWidth, newHeight);
                    rectangles.Add(currRect);
                }
            }
            // Build the Image
            var stitchedImage = new Bitmap(totalWidth, totalHeight);
            // Get all Screenshots and stitch them together
            var previous = Rectangle.Empty;
            foreach (var rectangle in rectangles)
            {
                // Calculate the scrolling (if needed)
                if (previous != Rectangle.Empty)
                {
                    var xDiff = rectangle.Right - previous.Right;
                    var yDiff = rectangle.Bottom - previous.Bottom;
                    // Scroll
                    ((QA.IJavaScriptExecutor)this.wd).ExecuteScript(String.Format("window.scrollBy({0}, {1})", xDiff, yDiff));
                }
                // Take Screenshot
                var screenshot = ((QA.ITakesScreenshot)wd).GetScreenshot();
                // Build an Image out of the Screenshot
                var screenshotImage = ScreenshotToImage(screenshot);
                // Calculate the source Rectangle
                var sourceRectangle = new Rectangle(viewportWidth - rectangle.Width, viewportHeight - rectangle.Height, rectangle.Width, rectangle.Height);
                // Copy the Image
                using (var graphics = Graphics.FromImage(stitchedImage))
                {
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = PixelOffsetMode.Half;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                    graphics.DrawImage(screenshotImage, rectangle, sourceRectangle, GraphicsUnit.Pixel);
                }
                // Set the Previous Rectangle
                previous = rectangle;
            }
    ((QA.IJavaScriptExecutor)this.wd).ExecuteScript("window.scrollTo(0, 0)");
            return stitchedImage;
        }

        private static Image ScreenshotToImage(QA.Screenshot screenshot)
        {
            Image screenshotImage;
            using (var memStream = new MemoryStream(screenshot.AsByteArray))
            {
                screenshotImage = Image.FromStream(memStream);
            }
            return screenshotImage;
        }
        /// <summary>
        /// Move horizontal scroll bar to right for the page
        /// </summary>
        public void PageScrollToRight()
        {
            var js = "document.documentElement.scrollLeft=10000";
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Chrome:
                    {
                        ((QA.Chrome.ChromeDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Firefox:
                    {
                        ((QA.Firefox.FirefoxDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Safari:
                    {
                        ((QA.Safari.SafariDriver)wd).ExecuteScript(js, null);
                    }; break;
                default:
                    {
                        ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                    }; break;
            }
        }


        /// <summary>
        /// Move vertical scroll bar to bottom for an element
        /// </summary>
        /// <param name="element"></param>
        public void ElementScrollToBottom(QA.IWebElement element)
        {
            string id = element.GetAttribute("id");
            string name = element.GetAttribute("name");
            var js = "";
            if (!string.IsNullOrWhiteSpace(id))
            {
                js = "document.getElementById('" + id + "').scrollTop=10000";
            }
            else if (!string.IsNullOrWhiteSpace(name))
            {
                js = "document.getElementsByName('" + name + "')[0].scrollTop=10000";
            }
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Chrome:
                    {
                        ((QA.Chrome.ChromeDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Firefox:
                    {
                        ((QA.Firefox.FirefoxDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Safari:
                    {
                        ((QA.Safari.SafariDriver)wd).ExecuteScript(js, null);
                    }; break;
                default:
                    {
                        ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                    }; break;
            }
        }
        /// <summary>
        /// Get a screen shot of the current window
        /// </summary>
        /// <param name="savePath"></param>
        public QA.Screenshot GetScreenshot()
        {
            QA.Screenshot theScreenshot = null;
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        theScreenshot = ((QA.IE.InternetExplorerDriver)wd).GetScreenshot();
                    }; break;
                case Browsers.Chrome:
                    {
                        theScreenshot = ((QA.Chrome.ChromeDriver)wd).GetScreenshot();
                    }; break;
                case Browsers.Firefox:
                    {
                        theScreenshot = ((QA.Firefox.FirefoxDriver)wd).GetScreenshot();
                    }; break;
                case Browsers.Safari:
                    {
                        theScreenshot = ((QA.Safari.SafariDriver)wd).GetScreenshot();
                    }; break;
                default:
                    {
                        theScreenshot = ((QA.IE.InternetExplorerDriver)wd).GetScreenshot();
                    }; break;
            }
            return theScreenshot;
        }

        /// <summary>
        /// Get a screen shot of the current window
        /// </summary>
        /// <param name="savePath"></param>
        public void TakeScreenshot(string savePath)
        {
            QA.Screenshot theScreenshot = null;
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        theScreenshot = ((QA.IE.InternetExplorerDriver)wd).GetScreenshot();
                    }; break;
                case Browsers.Chrome:
                    {
                        theScreenshot = ((QA.Chrome.ChromeDriver)wd).GetScreenshot();
                    }; break;
                case Browsers.Firefox:
                    {
                        theScreenshot = ((QA.Firefox.FirefoxDriver)wd).GetScreenshot();
                    }; break;
                case Browsers.Safari:
                    {
                        theScreenshot = ((QA.Safari.SafariDriver)wd).GetScreenshot();
                    }; break;
                default:
                    {
                        theScreenshot = ((QA.IE.InternetExplorerDriver)wd).GetScreenshot();
                    }; break;
            }
            if (theScreenshot != null)
            {
                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        theScreenshot.SaveAsFile(savePath, QA.ScreenshotImageFormat.Jpeg);
                        return;
                    }
                    catch (Exception ex)
                    {
                        Commons.Wait(new Random().Next(3, 7));
                    }
                }
            }
        }

        /// <summary>
        /// Find the element of a specified id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public QA.IWebElement FindElementById(string id)
        {
            try
            {
                QA.IWebElement theElement = null;
                theElement = (QA.IWebElement)wd.FindElement(QA.By.Id(id));
                return theElement;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Find the element of a specified name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public QA.IWebElement FindElementByName(string name)
        {
            QA.IWebElement theElement = null;
            theElement = (QA.IWebElement)wd.FindElement(QA.By.Name(name));
            return theElement;
        }

        /// <summary>
        /// Find the element by xpath
        /// </summary>
        /// <param name="xpath"></param>
        /// <returns></returns>
        public QA.IWebElement FindElementByXPath(string xpath)
        {
            QA.IWebElement theElement = null;
            try
            {
                theElement = (QA.IWebElement)wd.FindElement(QA.By.XPath(xpath));
            }
            catch (Exception)
            {
                return null;
            }
            return theElement;
        }

        public QA.IWebElement FindElementByLinkText(string text)
        {
            QA.IWebElement theElement = null;
            try
            {
                theElement = wd.FindElement(QA.By.LinkText(text));
            }
            catch { }
            return theElement;
        }

        public IList<QA.IWebElement> FindElementsByLinkText(string text)
        {
            IList<QA.IWebElement> theElement = null;
            theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.LinkText(text));
            return theElement;
        }

        public IList<QA.IWebElement> FindElementsByPartialLinkText(string text)
        {
            IList<QA.IWebElement> theElement = null;
            theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.PartialLinkText(text));
            return theElement;
        }

        public IList<QA.IWebElement> FindElementsByClassName(string clsName)
        {
            IList<QA.IWebElement> theElement = null;
            try
            {

                theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.ClassName(clsName));
            }
            catch (Exception)
            {
            }
            return theElement;
        }

        public IList<QA.IWebElement> FindElementsByTagName(string tagName)
        {
            IList<QA.IWebElement> theElement = null;
            theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.TagName(tagName));
            return theElement;
        }

        public IList<QA.IWebElement> FindElementsByCssSelector(string css)
        {
            IList<QA.IWebElement> theElement = null;
            theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.CssSelector(css));
            return theElement;
        }

        public IList<QA.IWebElement> FindElementsByXPathName(string xpath)
        {
            IList<QA.IWebElement> theElement = null;
            theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.XPath(xpath));
            return theElement;
        }

        public string GetSource()
        {
            try
            {
                return this.wd.PageSource;
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// Executes javascript
        /// </summary>
        /// <param name="js"></param>
        public void ExecuteJS(string js)
        {
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Chrome:
                    {
                        ((QA.Chrome.ChromeDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Firefox:
                    {
                        ((QA.Firefox.FirefoxDriver)wd).ExecuteScript(js, null);
                    }; break;
                case Browsers.Safari:
                    {
                        ((QA.Safari.SafariDriver)wd).ExecuteScript(js, null);
                    }; break;
                default:
                    {
                        ((QA.IE.InternetExplorerDriver)wd).ExecuteScript(js, null);
                    }; break;
            }
        }

        public void ClickElement(QA.IWebElement element)
        {
            (new QA.Interactions.Actions(wd)).Click(element).Perform();
        }

        public void DoubleClickElement(QA.IWebElement element)
        {
            (new QA.Interactions.Actions(wd)).DoubleClick(element).Perform();
        }

        public void ClickAndHoldOnElement(QA.IWebElement element)
        {
            (new QA.Interactions.Actions(wd)).ClickAndHold(element).Perform();
        }

        public void ContextClickOnElement(QA.IWebElement element)
        {
            (new QA.Interactions.Actions(wd)).ContextClick(element).Perform();
        }

        public void DragAndDropElement(QA.IWebElement source, QA.IWebElement target)
        {
            (new QA.Interactions.Actions(wd)).DragAndDrop(source, target).Perform();
        }

        public void SendKeysToElement(QA.IWebElement element, string text)
        {
            (new QA.Interactions.Actions(wd)).SendKeys(element, text).Perform();
        }



        public static void KillProcessAndChildren(int pid)
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("Select * From Win32_Process Where ParentProcessID=" + pid);
                ManagementObjectCollection moc = searcher.Get();
                foreach (ManagementObject mo in moc)
                {
                    KillProcessAndChildren(Convert.ToInt32(mo["ProcessID"]));
                }
            }
            catch { }

            try
            {
                System.Diagnostics.Process proc = System.Diagnostics.Process.GetProcessById(pid);
                proc.Kill();
            }
            catch (Exception)
            {

            }

            try
            {
                closeProcPID(pid);
            }
            catch (Exception)
            {
            }
        }

        public static bool closeProcPID(int pid)
        {
            bool result = false;
            System.Collections.ArrayList procList = new System.Collections.ArrayList();
            try
            {
                foreach (System.Diagnostics.Process thisProc in System.Diagnostics.Process.GetProcesses())
                {
                    try
                    {
                        if (thisProc.Id.Equals(pid))
                        {
                            thisProc.Kill();
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return result;
        }
        /// <summary>
        /// Quit this server, close all windows associated to it
        /// </summary>
        public void Quit()
        {
            for (int i = 0; i < 4; i++)
            {
                try
                {
                    wd.Quit();
                    try
                    {
                        KillProcessAndChildren(this.ProcessID);
                    }
                    catch (Exception)
                    {
                    }
                    break;
                }
                catch (Exception)
                {
                    Commons.Wait(1);
                }
            }
        }
        #endregion
    }

    public enum Browsers
    {
        IE,
        Firefox,
        Chrome,
        Safari,
        PhantomJS
    }

}
