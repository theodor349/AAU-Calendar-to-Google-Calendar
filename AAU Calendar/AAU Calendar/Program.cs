using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using HtmlAgilityPack;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace AAU_Calendar
{
    class Program
    {
        static List<Lector> Lectors = new List<Lector>();
        static bool isDone = false;

        static void Main(string[] args)
        {
            Process();

            // Wait for program to finish before terminating
            while (!isDone)
            {
                Thread.Sleep(1000);
            }
        }

        static async void Process()
        {
            // Find the week to retrieve data from
            int week = GetWeekToGet();
            if (week == -1)
            {
                isDone = true;
                return;
            }
            // What year should we retrieve data from
            int year = GetYear(week);

            // Get the Calendar
            var task = GetCallendar(week);
            try
            {
                await task;
            }
            catch(Exception e)
            {
                isDone = true;
                return;
            }

            // Post the lectors to Google Calendar
            task = GoogleCalendar.PostEvents(Lectors, year, FirstDateOfWeekISO8601(year, week));
            await task;

            // Write to file, what week has been retrieved
            WriteWeekToFile(week);
            // When done let the main thread terminate the program
            isDone = true;
        }

        private static int GetWeekToGet()
        {
            int currentWeek = GetIso8601WeekOfYear(DateTime.Now);
            int lastWeek = ReadWeekFromFile();

            int dif = lastWeek - GetIso8601WeekOfYear(DateTime.Now);
            bool isWeekday = (DateTime.Now - FirstDateOfWeekISO8601(DateTime.Now.Year, GetIso8601WeekOfYear(DateTime.Now))).Days < 5;

            // If the calendar haven't been update for more than a week run
            if (dif < 0)
                return currentWeek;
            // If it is not weekend or we are up to date
            else if (isWeekday || dif > 0)
                return -1;
            // It is weekend and we have yet to retrieve next weeks calendar
            else
                return lastWeek + 1 > GetWeeksInYear(DateTime.Now.Year) ? 1 : lastWeek + 1;
        }

        private static int GetYear(int week)
        {
            int cWeek = GetIso8601WeekOfYear(DateTime.Now);
            // Looking into next year?
            if (cWeek > week)
                return DateTime.Now.Year + 1;
            return DateTime.Now.Year;

        }

        private static async Task GetCallendar(int week)
        {
            if (week == -1)
                return;

            // Get URL to retrieve calendar from
            string url = ReadCalUrlFromFile();

            var htmlDoc = new HtmlDocument();
            try
            {
                // Retrieve the HTML page
                var httpClient = new HttpClient();
                var html = await httpClient.GetStringAsync(url);
                htmlDoc.LoadHtml(html);
            }
            catch(Exception e)
            {
                Console.WriteLine("You need to supply a valid URL in the properties.txt file");
            }

            // Get the schedule part
            var schedule = htmlDoc.DocumentNode.Descendants("div")
                .Where(node => node.GetAttributeValue("id", "")
                .Equals("schedule")).ToList();

            // Retrieve the lectors from the given week
            var days = GetDays(schedule, "week" + week.ToString());
            var events = GetLectors(days);

            // Add it to the static lectors variable
            foreach (var l in events)
            {
                Lectors.Add(l);
            }
        }

        private static List<HtmlNode> GetDays(List<HtmlNode> schedule, string week)
        {
            bool isCorrectWeek = false;
            var days = new List<HtmlNode>();
            // Go through each element and only take the days we are interested in
            foreach (var line in schedule[0].Descendants())
            {
                // Is it the correct week and a day?
                if (line.Attributes.Count > 0 && line.Attributes[0].Value.Equals("day") && isCorrectWeek)
                {
                    days.Add(line);
                }

                // If the element has the week we are looking for
                if((line.Attributes["id"] == null ? " " : line.Attributes["id"].Value).Equals(week))
                {
                    isCorrectWeek = true;
                }
                // Is this a week?
                else if ((line.Attributes["class"] == null ? " " : line.Attributes["class"].Value).Equals("week"))
                {
                    // Have we retrieved the days we need?
                    if(isCorrectWeek)
                        break;
                }
            }

            return days;
        }

        private static List<Lector> GetLectors(List<HtmlNode> days)
        {
            var events = new List<Lector>();

            // Days
            for (int i = 0; i < days.Count; i++)
            {
                var es = days[i].Descendants("div")
                    .Where(node => node.GetAttributeValue("class", "")
                    .Equals("event")).ToList();

                // Events (Lectors)
                foreach (var e in es)
                {
                    var myEvent = new Lector();
                    myEvent.Class = GetEventTitle(e);
                    myEvent.Teacher = GetEventValue(e, "teacher");
                    myEvent.Location = GetEventValue(e, "location");
                    myEvent.Note = GetEventValue(e, "note");
                    var time = GetEventValue(e, "time").Split(" - ");
                    myEvent.Start = time[0].Substring(5);
                    myEvent.End = time[1];
                    myEvent.Day = i;

                    events.Add(myEvent);
                }
            }

            return events;
        }

        private static string GetEventTitle(HtmlNode n)
        {
            var titles = n.Descendants("a").ToList();

            return titles[0].InnerHtml;
        }

        private static string GetEventValue(HtmlNode n, string attribute)
        {
            var attributes = n.Descendants("div")
                .Where(node => node.GetAttributeValue("class", "")
                .Equals(attribute)).ToList();

            return attributes[0].InnerHtml;
        }

        private static string ReadCalUrlFromFile()
        {
            string res = "";
            try
            {
                using (StreamReader sr = new StreamReader("properties.txt"))
                {
                    string line = sr.ReadLine();
                    res = line.Substring(12);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return res;
        }

        private static int ReadWeekFromFile()
        {
            int week = -1;
            try
            {
                using (StreamReader sr = new StreamReader("week.txt"))
                {
                    string line = sr.ReadLine();
                    try
                    {
                        week = int.Parse(line);
                    }
                    catch(Exception e)
                    {

                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return week;
        }

        private static void WriteWeekToFile(int week)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter("week.txt"))
                {
                    sw.WriteLine(week.ToString());
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static int GetWeeksInYear(int year)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime date1 = new DateTime(year, 12, 31);
            System.Globalization.Calendar cal = dfi.Calendar;
            return cal.GetWeekOfYear(date1, dfi.CalendarWeekRule,
                                                dfi.FirstDayOfWeek);
        }

        // This presumes that weeks start with Monday.
        // Week 1 is the 1st week of the year with a Thursday in it.
        public static int GetIso8601WeekOfYear(DateTime time)
        {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        public static DateTime FirstDateOfWeekISO8601(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            // Use first Thursday in January to get first week of the year as
            // it will never be in Week 52/53
            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;
            // As we're adding days to a date in Week 1,
            // we need to subtract 1 in order to get the right date for week #1
            if (firstWeek == 1)
            {
                weekNum -= 1;
            }

            // Using the first Thursday as starting week ensures that we are starting in the right year
            // then we add number of weeks multiplied with days
            var result = firstThursday.AddDays(weekNum * 7);

            // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
            return result.AddDays(-3);
        }
    }

    class Week
    {
        public int WeekNumber;

        public Week()
        {
        }

        public Week(int weekNumber)
        {
            WeekNumber = weekNumber;
        }
    }

    class Lector
    {
        public string Class;
        public string Teacher;
        public string Start;
        public string End;
        public string Location;
        public string Note;
        public int Day;

        public Lector()
        {
        }

        public Lector(string @class, string teacher, string start, string end, string location, string note, int day)
        {
            Class = @class;
            Teacher = teacher;
            Start = start;
            End = end;
            Location = location;
            Note = note;
            Day = day;
        }

        public void Print()
        {
            Console.WriteLine(Class);
            Console.WriteLine("Teacher: " + Teacher);
            Console.WriteLine("Time: " + Start + " - " + End);
            Console.WriteLine("Location: " + Location);
            Console.WriteLine("Note: " + Note);
        }
    }

}
