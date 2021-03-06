﻿using System;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Text;
using System.Reflection;

namespace Hull_Timetable_To_Outlook
{
    class Program
    {
        private static List<Lecture> timetable_lectures = new List<Lecture>();

        private static readonly string timetable_url = "https://timetable.hull.ac.uk/";
        private static string m_username = "";
        private static string m_password = "";

        private static IWebDriver driver;

        //Definitions
        private static DateTime week1 = new DateTime(2017, 8, 21);
        private static readonly string[] daysOfTheWeek = { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday" };

        static void Main(string[] args)
        {
            GrabUserDetails();

            try
            {
                //Checks for the geckodriver install
                //https://github.com/mozilla/geckodriver/releases
                //It needs to be downloaded and added to the PATH enviroment variables
                driver = new FirefoxDriver();
            }
            catch (System.Exception e)
            {
                //Prints a closes
                Console.WriteLine(e.Message);
                Console.ReadKey();
                return;
            }

            LoginToPage();
            var Elements = WorkThroughTable();
            GrabTimeTableData(Elements);
            driver.Quit();
            SyncToOutlook();
        }
        static void Restart()
        {
            var fileName = Assembly.GetExecutingAssembly().Location;
            System.Diagnostics.Process.Start(fileName);
            Environment.Exit(0);
        }

        static void GrabUserDetails()
        {
            Console.Write("Please enter your username: ");
            m_username = Console.ReadLine();

            Console.Write("Please enter your password: ");

            ConsoleKeyInfo consoleKey;
            do
            {
                consoleKey = Console.ReadKey(true);

                if (consoleKey.Key == ConsoleKey.Backspace)
                {
                    //Takes a value off the string
                    m_password = m_password.Substring(0, m_password.Length - 1);
                    Console.Write("\b");
                }
                else
                {
                    //Obssurces the key
                    m_password += consoleKey.KeyChar;
                    Console.Write("*");
                }
            }
            while (consoleKey.Key != ConsoleKey.Enter);
            Console.WriteLine();
        }
        static void LoginToPage()
        {
            driver.Navigate().GoToUrl(timetable_url);

            //Inserts username
            FillTextBox("tUserName", m_username);
            FillTextBox("tPassword", m_password);

            //Login click
            ClickElement("bLogin");

            try
            {
                // Selects value
                ClickElement("LinkBtn_studentMyTimetableNEW");

                //Shows timetable
                ClickElement("bGetTimetable");
            }
            catch (System.Exception)
            {
                driver.Quit();
                Console.WriteLine("\n ## Login deatils were not correct try again! ##");
                Console.WriteLine("Press enter to restart...");
                Console.ReadKey();
                Restart();
            }
            
        }

        static void SyncToOutlook()
        {
            foreach (var lecture in timetable_lectures)
            {
                TimeSpan startTime = TimeSpan.FromHours(lecture.StartTime);
                TimeSpan endTime = TimeSpan.FromHours(lecture.EndTime);

                //Grabs how many days need to be added to the week start
                int weekIndex = Array.FindIndex(daysOfTheWeek, w => w == lecture.WeekDay);

                //Loops round all the time periods and breaks them down into individual weeks
                foreach (var week in lecture.TimePeriod)
                {
                    for (int i = week.StartWeek; i <= week.EndWeek; i++)
                    {
                        DateTime date = WeekNumberToDate(i);

                        //Moves the date from the start of the week
                        date = date.AddDays(weekIndex);

                        DateTime start = date.Add(startTime);
                        DateTime end = date.Add(endTime);

                        string subject = lecture.ModuleTitle + " " + lecture.LecturerName;
                        string location = lecture.RoomName;

                        CreateAppoitment(start, end, subject, location);
                    }
                }
            }
            Console.WriteLine("DONE!");
        }
        static void CreateAppoitment(DateTime start, DateTime end, string subject, string location)
        {
            AppointmentItem appointment = null;
            Application application = new Application();

            var mapiNamespace = application.GetNamespace("MAPI");
            var CalendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            var appointmentItems = CalendarFolder.Items;

            try
            {
                appointment = (AppointmentItem)application.CreateItem(OlItemType.olAppointmentItem);
                appointment.Start = start;
                appointment.End = end;
                appointment.Subject = subject;
                appointment.Location = location;
                appointment.BusyStatus = OlBusyStatus.olOutOfOffice;

                //Checks for exisiting appointments
                //TODO: This stops duplicates from being created, but what if another appointment was a confict that was unrelated to uni?
                if (appointment.Conflicts == null)
                {
                    appointment.Save();
                }
            }
            finally
            {
                if (appointment != null) Marshal.ReleaseComObject(appointment);
            }
        }

        static List<Element> WorkThroughTable()
        {
            var Elements = new List<Element>();

            var table = driver.FindElement(By.XPath("/html/body/table[2]/tbody"));
            var tableRows = new List<IWebElement>(table.FindElements(By.XPath("*")));

            //Removes headings
            tableRows.RemoveAt(0);

            //Removes the weekend
            tableRows.RemoveAt(tableRows.Count - 2);
            tableRows.RemoveAt(tableRows.Count - 1);

            //Day calculation
            var dayLabels = new List<IWebElement>(table.FindElements(By.ClassName("row-label-one")));
            dayLabels.RemoveAt(dayLabels.Count - 2);
            dayLabels.RemoveAt(dayLabels.Count - 1);

            //This section of code creates a list that matches the length of the row list
            //Each position is the rows corresponding day of the week
            List<string> dayCounter = new List<string>();
            int currentDay = 0;
            foreach (var label in dayLabels)
            {
                //Grabs how many rows the day covers and adds it that many times to the list
                int span = int.Parse(label.GetAttribute("rowspan"));

                for (int i = 0; i < span; i++)
                {
                    dayCounter.Add(daysOfTheWeek[currentDay]);
                }

                currentDay++;
            }

            //This goes thought and grabs the cells and calculats the time and day that event is occuring
            int dayCount = 0;
            foreach (IWebElement row in tableRows)
            {
                var all_cells = row.FindElements(By.XPath("*"));

                //What time the element is in
                int position = 0;

                string weekDay = dayCounter[dayCount];

                // Check if the values have a class
                foreach (var cell in all_cells)
                {
                    //Checks if they're separate cell value 
                    if (cell.GetAttribute("class") == "cell-border" ||
                        cell.GetAttribute("class") == "object-cell-border")
                    {
                        double startTime = 9 + (position * 0.25);

                        //Moves along the position value
                        string span = cell.GetAttribute("colspan");
                        if (span != null)
                        {
                            position += int.Parse(span);
                        }
                        else
                        {
                            position++;
                        }

                        double endTime = 9 + (position * 0.25) - 0.25;

                        //Only adds if they're values to deal with
                        if (cell.Text.Trim() != "")
                        {
                            Elements.Add(new Element(cell.Text, startTime, endTime, weekDay));
                        }
                    }
                }
                dayCount++;
            }

            return Elements;
        }
        static void GrabTimeTableData(List<Element> Elements)
        {
            foreach (var element in Elements)
            {
                string moduleCode;
                string moduleTitle;
                string lecturerName;
                string roomName;

                var lines = element.Text.Split('\n');

                //First line contains Module code and lecture ID
                var firstLineParts = lines[0].Split(' ');
                moduleCode = firstLineParts[0];

                //Second line contains Module name and location
                if (!lines[1].Contains("[CANCELLED]"))
                {
                    var secondLineParts = lines[1].Split(' ');

                    //Grab room name
                    roomName = secondLineParts[secondLineParts.Length - 1];

                    //Use the length of the room name to trim off the row to obtain the module name
                    int startIndex = lines[1].Length - roomName.Length;
                    moduleTitle = lines[1].Remove(startIndex, roomName.Length - 1);

                }
                else { continue; } //Ignore is canceled

                //Third line contains Lecturer name
                string regex_name = @"[a-zA-Z]+, +[a-zA-Z]+";
                lecturerName = Regex.Match(lines[2], regex_name).Value;

                //Data range is obtained by a Regex that looks for "<" and ">" that encase a date range
                string regex_dataRange = @"(([0-9]+\-[0-9]+)|([0-9]+))(\, (([0-9]+\-[0-9]+)|([0-9]+)))*";
                string weekRanges = Regex.Match(lines[2], regex_dataRange).Value;

                List<WeekRange> weekRangeList = new List<WeekRange>();
                if (weekRanges != "")
                {
                    //If it's a multi week
                    if (weekRanges.Contains(","))
                    {
                        var parts = weekRanges.Split(',');

                        foreach (var item in parts)
                        {
                            weekRangeList.Add(ParseWeekRange(item));
                        }
                    }
                    //If it's a single week
                    else
                    {
                        weekRangeList.Add(ParseWeekRange(weekRanges));
                    }
                }
                else
                {
                    continue;
                } // Skip this value

                //Adds the completed object
                Lecture lecture = new Lecture(
                    element.StartTime,
                    element.EndTime,
                    element.weekDay,
                    moduleCode.Trim(),
                    moduleTitle.Trim(),
                    lecturerName.Trim(),
                    roomName.Trim(),
                    weekRangeList);

                timetable_lectures.Add(lecture);
            }
        }
        static WeekRange ParseWeekRange(string weekRanges)
        {
            //Possibilities for ranges
            // Multi range e.g    "<5-10, 12-13>"
            if (weekRanges.Contains("-") && weekRanges.Contains(","))
            {
                var weekParts = weekRanges.Split(',');

                foreach (var part in weekParts)
                {
                    var smallerParts = part.Split('-');

                    int start = int.Parse(smallerParts[0].Trim());
                    int end = int.Parse(smallerParts[1].Trim());

                    return new WeekRange(start, end);
                }
            }
            //Single range e.g. "<5-10>"
            else if (weekRanges.Contains("-"))
            {
                var parts = weekRanges.Split('-');

                int start = int.Parse(parts[0].Trim());
                int end = int.Parse(parts[1].Trim());

                return new WeekRange(start, end);
            }
            //Single week e.g. "<5>"
            else
            {
                return new WeekRange(int.Parse(weekRanges), int.Parse(weekRanges));
            }

            return null;
        }
        static DateTime WeekNumberToDate(int weekNumber)
        {
            return week1.AddDays(7 * weekNumber);
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

        static string MD5_Hash(string input_str)
        {
            byte[] input = System.Text.Encoding.ASCII.GetBytes(input_str);
            byte[] output = null;
            using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create())
            {
                output = md5.ComputeHash(input);
            }

            var sb = new StringBuilder();
            for (int i = 0; i < output.Length; i++)
            {
                sb.Append(output[i].ToString("X2"));
            }
            return sb.ToString();
        }
    }

    class Lecture
    {
        public Lecture(double pStartTime, double pEndTime, string pWeekDay, string pModuleCode, string pModuleTitle, string pLecturerName, string pRoomName, List<WeekRange> pTimePeriod)
        {
            StartTime = pStartTime;
            EndTime = pEndTime;
            WeekDay = pWeekDay;
            ModuleCode = pModuleCode;
            ModuleTitle = pModuleTitle;
            LecturerName = pLecturerName;
            RoomName = pRoomName;
            TimePeriod = pTimePeriod;
        }

        public double StartTime;
        public double EndTime;
        public string WeekDay;
        public string ModuleCode { get; }
        public string ModuleTitle { get; }
        public string LecturerName { get; }
        public string RoomName { get; }
        public List<WeekRange> TimePeriod { get; }
    }

    class WeekRange
    {
        public WeekRange(int pStartWeek, int pEndWeek)
        {
            StartWeek = pStartWeek;
            EndWeek = pEndWeek;
        }

        public int StartWeek { get; }
        public int EndWeek { get; }
    }

    class Element
    {
        public Element(string pElement, double pStartTime, double pEndTime, string pWeekDay)
        {
            Text = pElement;
            StartTime = pStartTime;
            EndTime = pEndTime;
            weekDay = pWeekDay;
        }

        public string Text { get; }
        public double StartTime { get; }
        public double EndTime { get; }
        public string weekDay { get; }
    }
}
