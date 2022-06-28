using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutoClick
{
    class Program
    {
        #region Property
        public static string LoginAccount { get; set; } = "";
        public static string LoginPassword { get; set; } = "";

        public static bool IsLogin = false;

        public static IWebDriver Driver
        {
            get; set;
        }
        #endregion

        #region Memberfunction
        static void Main(string[] args)
        {

            string downloadPath = @"D:\SeleniumDownload";
            string loginUrl = "http://trade03.cathaysec.com.tw/EIPWeb/login.jsp";

            var account = ConfigurationSettings.AppSettings["Account"];
            var password = ConfigurationSettings.AppSettings["Password"];
            if(account is null) 
            {
                Console.WriteLine("Cant read Account from config");
                return;
            }
            else 
            {
                LoginAccount = account;
            }
            if (password is null)
            {
                Console.WriteLine("Cant read Password from config");
                return;
            }
            else 
            {
                LoginPassword = password;
            }


            DirectoryInfo di = new DirectoryInfo(downloadPath);

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                dir.Delete(true);
            }

            var chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.default_directory", downloadPath);
            //chromeOptions.AddArguments("headless");
            var Cookie = new List<string>();
            Driver = new ChromeDriver(chromeOptions);
            Driver.Navigate().GoToUrl(loginUrl);
            Thread.Sleep(8000);
            var idInput = Driver.FindElement(By.Name("id"));
            idInput.SendKeys(LoginAccount);
            var pwInput = Driver.FindElement(By.Name("pass"));
            pwInput.SendKeys(LoginPassword);
            var submit = Driver.FindElement(By.CssSelector("input[value='送出']"));
            submit.Click();
            Thread.Sleep(8000);


            var rootIframe = Driver.FindElement(By.TagName("iframe"));
            Driver.SwitchTo().Frame(rootIframe);
            //var dt = Driver.PageSource;
            var mainFrame = Driver.FindElement(By.CssSelector("frame[name='main']"));
            Driver.SwitchTo().Frame(mainFrame);
            //var dt = Driver.PageSource;
            var mainLeft = Driver.FindElement(By.CssSelector("frame[name='left']"));
            Driver.SwitchTo().Frame(mainLeft);
            //var dt = Driver.PageSource;

            Driver.FindElement(By.XPath("//a[text() = '法令遵循區']")).Click();

            Thread.Sleep(5000);


            Driver.Navigate().GoToUrl("http://secsvr325/DMSWeb/frame/eipmain_index.jsp");
            Thread.Sleep(2000);

            //到法令遵循專區
            var FinalleftFrame = Driver.FindElement(By.CssSelector("frame[name='left']"));
            var FinalMainFrame = Driver.FindElement(By.CssSelector("frame[name='mainFrame']"));
            Driver.SwitchTo().Frame(FinalleftFrame);

            //點個人點閱紀錄查詢
            string excuteJsResult = string.Empty;
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
            excuteJsResult = (string)js.ExecuteScript("javascript:submitForm2('UC01020B','default',0,'123')");
            Thread.Sleep(5000);
            //回到root
            Driver.SwitchTo().DefaultContent();

            //切換到主要呈現資料頁面
            Driver.SwitchTo().Frame(FinalMainFrame);

            //搜尋
            excuteJsResult = (string)js.ExecuteScript("javascript: linkTo('UC01020B', 'GetEIPCheckZoneViewB')");
            Thread.Sleep(5000);


            var table = Driver.FindElement(By.TagName("table"));
            var allLink = table.FindElements(By.TagName("a"));
            foreach (var item in allLink)
            {
                if (!string.IsNullOrEmpty(item.Text))
                {
                    if (!item.Text.Contains("搜尋"))
                    {
                        Console.WriteLine($"Click {item.Text}");
                        item.Click();
                    }
                }
            }

           
            Console.WriteLine("Done");
            Driver.Close();
            Driver.Dispose();
            Driver.Quit();

        }
        #endregion

    }
}
