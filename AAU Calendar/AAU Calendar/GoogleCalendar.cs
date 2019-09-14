using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace AAU_Calendar
{
    internal class GoogleCalendar
    {
        private static string[] Scopes = { CalendarService.Scope.Calendar };
        private static string ApplicationName = "AAU Calendar";

        public static async Task PostEvents(List<Lector> lectors, int year, DateTime startDay)
        {
            var credential = GetCredentials();
            string calendarId = ReadCalIdFromFile();

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Lectors
            foreach (var l in lectors)
            {
                PostEvent(l, service, year, startDay, calendarId);
            }
            Console.WriteLine("Events Add: " + lectors.Count);
        }

        private static void PostEvent(Lector l, CalendarService service, int year, DateTime startDay, string calendarId)
        {
            Console.WriteLine("Posting Event");
            var ev = new Event();

            var timeS = l.Start.Split(':');
            var timeE = l.End.Split(':');

            var date = startDay.AddDays(l.Day);

            // Format the date and time of the lector
            EventDateTime start = new EventDateTime();
            start.DateTime = new DateTime(year, date.Month, date.Day, int.Parse(timeS[0]), int.Parse(timeS[1]), 0);
            EventDateTime end = new EventDateTime();
            end.DateTime = new DateTime(year, date.Month, date.Day, int.Parse(timeE[0]), int.Parse(timeE[1]), 0);

            // Set event data
            ev.Start = start;
            ev.End = end;
            ev.Location = l.Location;
            ev.Summary = l.Class;
            ev.Description = l.Teacher + "\n" + l.Note;

            try
            {
                // Post event
                service.Events.Insert(ev, calendarId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("You need to supply a valid CalendarID in the properties.txt file");
            }
        }

        private static UserCredential GetCredentials()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }
        private static string ReadCalIdFromFile()
        {
            string res = "";
            try
            {
                using (StreamReader sr = new StreamReader("properties.txt"))
                {
                    string line = sr.ReadLine();
                    line = sr.ReadLine();
                    res = line.Substring(11);
                    while (res[0].Equals(' '))
                        res = res.Remove(0, 1);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return res;
        }
    }
}