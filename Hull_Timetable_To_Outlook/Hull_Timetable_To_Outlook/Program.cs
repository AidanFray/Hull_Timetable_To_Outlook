using System;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Collections.Generic;

namespace Hull_Timetable_To_Outlook
{
    class Program
    {
        private static readonly string timetable_url = "https://timetable.hull.ac.uk/";

        private static string m_username = "";
        private static string m_password = "";


        private static IWebDriver driver = new FirefoxDriver();

        static void Main(string[] args)
        {
            GrabUserDetails();
            LoginToPage();
            GrabTimeTableData();

            driver.Quit();
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
            driver.Navigate().GoToUrl(timetable_url);

            //Inserts username
            FillTextBox("tUserName", m_username);
            FillTextBox("tPassword", m_password);

            //Login click
            ClickElement("bLogin");

            // Selects value
            ClickElement("LinkBtn_studentMyTimetableNEW");
            
            //Shows timetable
            ClickElement("bGetTimetable");
        }

        static void GrabTimeTableData()
        {
            var TimetableElements = driver.FindElements(By.ClassName("object-cell-border"));

            Console.WriteLine(TimetableElements[0].Text);
        }

        static void FillTextBox(string form_id, string value)
        {
            IWebElement query = driver.FindElement(By.Id(form_id));
            query.SendKeys(value);
        }
        static void ClickElement(string form_id)
        {
            IWebElement query = driver.FindElement(By.Id(form_id));
            query.Click();
        }
    }

    class Lecture
    {
        public Lecture(int pModuleCode, string pModuleTitle, string pLecturerName, string pRoomName, List<WeekRange> pTimePeriod)
        {
            ModuleCode = pModuleCode;
            ModuleTitle = pModuleTitle;
            LecturerName = pLecturerName;
            RoomName = pRoomName;
            TimePeriod = pTimePeriod;
        }

        int ModuleCode { get; }
        string ModuleTitle { get; }
        string LecturerName { get; }
        string RoomName { get; }
        List<WeekRange> TimePeriod { get; }
    }

    struct WeekRange
    {
        int StartWeek;
        int EndWeek;
    }
}
