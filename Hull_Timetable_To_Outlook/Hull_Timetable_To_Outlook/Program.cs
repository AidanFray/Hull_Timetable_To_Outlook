using System;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;

namespace Hull_Timetable_To_Outlook
{
    class Program
    {
        private static List<Lecture> timetable_lectures = new List<Lecture>();

        private static readonly string timetable_url = "https://timetable.hull.ac.uk/";
        private static string m_username = "";
        private static string m_password = "";

        private static IWebDriver driver = new FirefoxDriver();
        
        //Definitions
        private static DateTime week1 = new DateTime(2017, 8, 21);
        private static readonly string[] daysOfTheWeek = { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday" };

        static void Main(string[] args)
        { 
            GrabUserDetails();
            LoginToPage();
            var Elements = WorkThroughTable();
            GrabTimeTableData(Elements);

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
                string regex_dataRange = @"<.+>";

                string weekRanges = Regex.Match(lines[0], regex_dataRange).Value;


                List<WeekRange> weekRangeList = new List<WeekRange>();
                if (weekRanges != "")
                {
                    // Removes the encasing "<" and ">"
                    weekRanges = weekRanges.Remove(0, 1);
                    weekRanges = weekRanges.Remove(weekRanges.Length - 1, 1);

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
                else { continue;  } // Skip this value

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
