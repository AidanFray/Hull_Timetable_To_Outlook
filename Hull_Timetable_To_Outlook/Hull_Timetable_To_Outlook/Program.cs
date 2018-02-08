using System;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;

namespace Hull_Timetable_To_Outlook
{
    class Program
    {
        private static readonly string timetable_url = "https://timetable.hull.ac.uk/";

        private static string m_username;
        private static string m_password;

        static void Main(string[] args)
        {
            GrabUserDetails();
            LoginToPage();
        }

        static void GrabUserDetails()
        {
            Console.Write("Please enter your username: ");
            m_username = Console.ReadLine();

            Console.Write("Please enter your password: ");
            m_password = Console.ReadLine();
        }
        
        static void LoginToPage()
        {
            IWebDriver driver = new FirefoxDriver();
            driver.Navigate().GoToUrl(timetable_url);

            //Inserts username
            FillTextBox("tUserName", m_username, driver);
            FillTextBox("tPassword", m_password, driver);

            //Login click
            ClickElement("bLogin", driver);

            // Selects value
            ClickElement("LinkBtn_studentMyTimetableNEW", driver);
            
            //Shows timetable
            ClickElement("bGetTimetable", driver);

            driver.Quit();
        }
        static void FillTextBox(string form_id, string value, IWebDriver driver)
        {
            IWebElement query = driver.FindElement(By.Id(form_id));
            query.SendKeys(value);
        }
        static void ClickElement(string form_id, IWebDriver driver)
        {
            IWebElement query = driver.FindElement(By.Id(form_id));
            query.Click();
        }
    }
}
