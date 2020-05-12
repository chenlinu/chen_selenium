
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Safari;
using System;
using System.Collections.Generic;
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
    public class SeleniumHelperCookie
    {
        public QA.IWebDriver wd = null;
        public Browsers browser = Browsers.Chrome;
        public QA.DriverService ds = null;
        private int _ProcessID = 0;
        public int ProcessID
        {
            get { return _ProcessID; }
        }
        public SeleniumHelperCookie(Browsers theBrowser, string cache_dir, bool loadImage)
        {
            this.browser = theBrowser;
            wd = InitWebDriver(cache_dir, loadImage);
        }
        public SeleniumHelperCookie(Browsers theBrowser, string cache_dir, bool loadImage, string useragent)
        {
            this.browser = theBrowser;
            wd = InitWebDriver(cache_dir, loadImage, useragent);
        }
        public SeleniumHelperCookie(Browsers theBrowser, string cache_dir)
        {
            this.browser = theBrowser;
            wd = InitWebDriver(cache_dir);
        }

        private QA.IWebDriver InitWebDriver(string cache_dir, bool loadImage = true, string useragent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36")
        {
            if (!Directory.Exists(cache_dir))
            {
                Directory.CreateDirectory(cache_dir);
            }
            QA.IWebDriver theDriver = null;
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        InternetExplorerDriverService driverService = InternetExplorerDriverService.CreateDefaultService(Application.StartupPath + "\\Drivers\\");
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;
                        QA.IE.InternetExplorerOptions _ieOptions = new QA.IE.InternetExplorerOptions();
                        _ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                        theDriver = new QA.IE.InternetExplorerDriver(driverService, _ieOptions);
                        _ProcessID = driverService.ProcessId;
                    }; break;
                case Browsers.Chrome:
                    {
                        ChromeDriverService driverService = ChromeDriverService.CreateDefaultService(Application.StartupPath + "\\Drivers\\");
                        driverService.Port = new Random().Next(1000, 2000);
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;
                        ChromeOptions options = new QA.Chrome.ChromeOptions();
                        options.AddArgument("--window-size=" + Screen.PrimaryScreen.WorkingArea.Width + "x" + Screen.PrimaryScreen.WorkingArea.Height);
                        options.AddArgument("--disable-gpu");
                        options.AddArgument("--disable-extensions");
                        options.AddArgument("--no-sandbox");
                        options.AddArgument("--disable-dev-shm-usage");
                        options.AddArgument("--disable-java");
                        options.AddArgument("--user-agent=" + useragent);
                        options.AddArgument(@"--user-data-dir=" + cache_dir);
                        if (loadImage == false)
                        {
                            options.AddUserProfilePreference("profile.managed_default_content_settings.images", 2);//不加载图片
                        }
                        theDriver = new QA.Chrome.ChromeDriver(driverService, options, TimeSpan.FromSeconds(240));
                        _ProcessID = driverService.ProcessId;
                    };
                    break;
                case Browsers.Firefox:
                    {
                        var driverService = FirefoxDriverService.CreateDefaultService(Application.StartupPath + "\\Drivers\\");
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;
                        FirefoxProfile profile = new FirefoxProfile();
                        if (loadImage == false)
                        {
                            profile.SetPreference("permissions.default.image", 2);
                            // 关掉flash  
                            profile.SetPreference("dom.ipc.plugins.enabled.libflashplayer.so", false);
                        }
                        FirefoxOptions options = new FirefoxOptions();
                        options.Profile = profile;
                        theDriver = new QA.Firefox.FirefoxDriver(driverService, options, TimeSpan.FromMinutes(240));
                        _ProcessID = driverService.ProcessId;
                    };
                    break;
                case Browsers.Safari:
                    {
                        SafariDriverService driverService = SafariDriverService.CreateDefaultService(Application.StartupPath + "\\Drivers\\");
                        driverService.HideCommandPromptWindow = true;
                        driverService.SuppressInitialDiagnosticInformation = true;
                        ds = driverService;
                        theDriver = new QA.Safari.SafariDriver(driverService);
                        _ProcessID = driverService.ProcessId;
                    };
                    break;

                default:
                    {
                        QA.IE.InternetExplorerOptions _ieOptions = new QA.IE.InternetExplorerOptions();
                        _ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                        theDriver = new QA.IE.InternetExplorerDriver(_ieOptions);
                    };
                    break;
            }
            //theDriver.Manage().Window.Maximize();
            return theDriver;
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
        public void ImplicitlyWaitRund(int max_seconds)
        {
            wd.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(new Random().Next(max_seconds));
        }

        /// <summary>
        /// Wait for the expected condition is satisfied, return immediately
        /// </summary>
        /// <param name="expectedCondition"></param>
        public void WaitForPage(string title, int seconds)
        {
            UI.WebDriverWait _wait = new UI.WebDriverWait(wd, TimeSpan.FromSeconds(seconds));
            _wait.Until((d) => { return d.Title.ToLower().StartsWith(title.ToLower()); });
            //to do
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="we"></param>
        public QA.IWebElement WaitForElement(string id, int seconds)
        {
            UI.WebDriverWait _wait = new UI.WebDriverWait(wd, TimeSpan.FromSeconds(seconds));
            _wait.Until((d) => { return SeleniumExtras.WaitHelpers.ExpectedConditions.PresenceOfAllElementsLocatedBy(QA.By.Id(id)); });
            return FindElementById(id);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="we"></param>
        public QA.IWebElement WaitForElementClass(string className, int seconds)
        {
            UI.WebDriverWait _wait = new UI.WebDriverWait(wd, TimeSpan.FromSeconds(seconds));
            _wait.Until((d) => { return SeleniumExtras.WaitHelpers.ExpectedConditions.PresenceOfAllElementsLocatedBy(QA.By.ClassName(className)); });
            return FindElementsByClassName(className).FirstOrDefault();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="we"></param>
        public QA.IWebElement WaitForElement(QA.By locator, int seconds)
        {
            UI.WebDriverWait _wait = new UI.WebDriverWait(wd, TimeSpan.FromSeconds(seconds));
            _wait.Until((d) => { return SeleniumExtras.WaitHelpers.ExpectedConditions.PresenceOfAllElementsLocatedBy(locator); });
            QA.IWebElement theElement = null;
            try
            {
                theElement = (QA.IWebElement)wd.FindElement(locator);
            }
            catch { }
            return theElement;
        }

        /// <summary>
        /// Load a new web page in current browser
        /// </summary>
        /// <param name="url"></param>
        public void GoToUrl(string url)
        {
            wd.Navigate().GoToUrl(url);
        }

        public void Refresh()
        {
            wd.Navigate().Refresh();
        }

        public void Back()
        {
            wd.Navigate().Back();
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
            return wd.Url;
        }

        /// <summary>
        /// Get page title of current browser window
        /// </summary>
        /// <returns></returns>
        public string GetPageTitle()
        {
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

        public void AddCookies(string cookies, string defaultDomain)
        {
            QA.ICookieJar ck = null;
            switch (this.browser)
            {
                case Browsers.IE:
                    {
                        ck = ((QA.IE.InternetExplorerDriver)wd).Manage().Cookies;

                    }; break;
                case Browsers.Chrome:
                    {
                        ck = ((QA.Chrome.ChromeDriver)wd).Manage().Cookies;
                    }; break;
                case Browsers.Firefox:
                    {
                        ck = ((QA.Firefox.FirefoxDriver)wd).Manage().Cookies;
                    }; break;
                case Browsers.Safari:
                    {
                        ck = ((QA.Safari.SafariDriver)wd).Manage().Cookies;
                    }; break;

                default:
                    {
                        ck = ((QA.IE.InternetExplorerDriver)wd).Manage().Cookies;
                    }; break;
            }

            List<QA.Cookie> list = ParseCookies(cookies, defaultDomain);
            foreach (QA.Cookie item in list)
            {
                ck.AddCookie(item);
            }
        }
        private List<QA.Cookie> ParseCookies(string cookie, string defaultDomain)
        {
            List<QA.Cookie> cc = new List<OpenQA.Selenium.Cookie>();
            try
            {
                Uri urI = new Uri(defaultDomain);
                defaultDomain = urI.Host.ToString();
                foreach (string item in cookie.Split(new string[] { ";", "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(item.Trim(), @"(?<key>[\s\S]*?)=(?<value>[\s\S]*?)$");
                    if (m.Success)
                    {
                        try
                        {
                            QA.Cookie c = new QA.Cookie(m.Groups["key"].Value.Trim(), m.Groups["value"].Value.Trim(), defaultDomain, "/", DateTime.Now.AddDays(1));
                            cc.Add(c);
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return cc;
        }
        /// <summary>
        /// Delete all cookies from the page
        /// </summary>
        public void DeleteAllCookies()
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
        public void PageScrollToTop()
        {
            var js = "document.documentElement.scrollTop=10";
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
        /// Move vertical scroll bar to bottom for the page
        /// </summary>
        public void PageScrollToBottom()
        {
            var js = "document.documentElement.scrollTop=210000";
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
        /// Move vertical scroll bar to bottom for the page
        /// </summary>
        public void PageScrollToBottom2()
        {
            try
            {
                int bottom = wd.Manage().Window.Size.Height;

                for (int i = 0; i < 8; i++)
                {
                    bottom += 200;
                    var js = "document.documentElement.scrollTop=" + bottom.ToString() + "";
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
                    Commons.Wait(2);
                }
            }
            catch (Exception ex)
            {
            }
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
                theScreenshot.SaveAsFile(savePath, QA.ScreenshotImageFormat.Jpeg);
            }
        }
        /// <summary>
        /// Get a screen shot of the current window
        /// </summary>
        /// <param name="savePath"></param>
        public Image TakeScreenshot()
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
                MemoryStream ms = new MemoryStream(theScreenshot.AsByteArray);
                Image image = Image.FromStream(ms);
                return image;
            }
            return null;
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


        public Bitmap CreateElementImage(QA.IWebElement el, int correct_left, int correct_right, int correct_top, int correct_bottom)
        {
            Image image = GetEntireScreenshot();
            Bitmap newBitmap = new Bitmap(el.Size.Width - correct_left - correct_right, el.Size.Height - correct_top - correct_bottom);
            Graphics g = null;
            try
            {
                g = Graphics.FromImage(newBitmap);
                //设置质量
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.DrawImage(image, new Rectangle(0, 0, newBitmap.Width, newBitmap.Height), el.Location.X + correct_left, el.Location.Y + correct_top, newBitmap.Width, newBitmap.Height, GraphicsUnit.Pixel);
            }
            finally
            {
                g.Dispose();
            }
            return newBitmap;
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
            catch { }
            return null;
        }

        /// <summary>
        /// Find the element of a specified name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public QA.IWebElement FindElementByName(string name)
        {
            try
            {
                QA.IWebElement theElement = null;
                theElement = (QA.IWebElement)wd.FindElement(QA.By.Name(name));
                return theElement;
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Find the element by xpath
        /// </summary>
        /// <param name="xpath"></param>
        /// <returns></returns>
        public QA.IWebElement FindElementByXPath(string xpath)
        {
            try
            {
                QA.IWebElement theElement = null;
                theElement = (QA.IWebElement)wd.FindElement(QA.By.XPath(xpath));
                return theElement;
            }
            catch { }
            return null;
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
            try
            {
                theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.LinkText(text));
            }
            catch { }
            return theElement;
        }

        public IList<QA.IWebElement> FindElementsByPartialLinkText(string text)
        {
            IList<QA.IWebElement> theElement = null;
            try
            {
                theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.PartialLinkText(text));
            }
            catch { }
            return theElement;
        }
        public QA.IWebElement FindElementByClassName(string clsName)
        {
            QA.IWebElement theElement = null;
            try
            {
                theElement = wd.FindElement(QA.By.ClassName(clsName));
            }
            catch { }
            return theElement;
        }
        public IList<QA.IWebElement> FindElementsByClassName(string clsName)
        {
            IList<QA.IWebElement> theElement = null;
            try
            {
                theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.ClassName(clsName));
            }
            catch { }
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
            try
            {
                theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.CssSelector(css));
            }
            catch { }
            return theElement;
        }

        public IList<QA.IWebElement> FindElementsByXPathName(string xpath)
        {
            IList<QA.IWebElement> theElement = null;
            try
            {
                theElement = (IList<QA.IWebElement>)wd.FindElements(QA.By.XPath(xpath));
            }
            catch { }
            return theElement;
        }

        public string GetSource()
        {
            return this.wd.PageSource;
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

        /**
       * 传入参数：父进程id
       * 功能：根据父进程id，杀死与之相关的进程树
       */
        public void KillProcessAndChildren(int pid)
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
        }

        /// <summary>
        /// Quit this server, close all windows associated to it
        /// </summary>
        public void Quit()
        {
            try
            {
                wd.Quit();
            }
            catch { }
            KillProcessAndChildren(ProcessID);
        }
        #endregion
    }
}