using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Spire.Presentation;
using Spire.Presentation.Collections;
using Path = System.IO.Path;

namespace CalendrierFamille
{
    static class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static readonly string[] _scopes = {CalendarService.Scope.CalendarReadonly};
        static readonly string _applicationName = "Calendrier Famille";

        static void Main(string[] args)
        {
            while (true)
            {
                var userProfile = Environment.GetEnvironmentVariable("USERPROFILE");
                Console.WriteLine("{0} Start update", DateTime.Now.ToString("s"));
                var presentation = new Spire.Presentation.Presentation();
                presentation.LoadFromFile(userProfile + @"\OneDrive\Famille\Menu de la semaine.pptx");

                var shape = (IAutoShape) presentation.Slides[3].Shapes[1];
                GetEvents(shape.TextFrame);
                var shape2 = (IAutoShape) presentation.Slides[3].Shapes[2];
                shape2.TextFrame.Paragraphs.Clear();
                shape2.TextFrame.Paragraphs.Append(new TextParagraph() {Text = "Dernière mise à jour: " + DateTime.Now.ToString("dddd le d MMMM yyyy @ H:mm").Translate()});
                presentation.SaveToFile(userProfile + @"\OneDrive\Famille\Menu de la semaine.pptx", FileFormat.Pptx2010);
                Console.WriteLine("{0} End update", DateTime.Now.ToString("s"));

                Console.WriteLine("{0} Waiting for next execution", DateTime.Now.ToString("s"));
                Thread.Sleep(60000);

                while (true)
                {
                    Console.Write(".");
                    var minute = DateTime.Now.Minute;
                    if (minute%15 == 0)
                    {
                        Console.WriteLine();
                        break;
                    }
                    Thread.Sleep(1000);
                }
            }
        }

        private static string GetEvents(ITextFrameProperties textFrame)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("calendrier-famille.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/calendar-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    _scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = _applicationName,
            });
            
            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 50;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
            request.TimeMax = DateTime.Now.AddDays(7);

            // List events.
            Events events = request.Execute();
            Console.WriteLine("Upcoming events:");

            var eventsByDay = from @event in events.Items
                let dateTime = @event.Start.DateTime
                where dateTime != null
                group @event by dateTime.Value.Date
                into day
                select new {Day = day.Key, Events = day};

            StringBuilder result = new StringBuilder();

            textFrame.Paragraphs.Clear();
            textFrame.Paragraphs.Append(new TextParagraph() { Text = "", BulletType = TextBulletType.None});
            var tp = textFrame.TextRange.Paragraph;
            tp.TextRanges.Clear();
            foreach (var currentDay in eventsByDay)
            {
                //result.AppendLine(currentDay.Day.ToString("dddd d MMMM").Translate());
                var textRange = new TextRange(string.Format("{0}\n", currentDay.Day.ToString("dddd d MMMM").Translate()));
                tp.TextRanges.Append(textRange);
                
                foreach (var @event in currentDay.Events)
                {
                    if (@event.Start.DateTime.HasValue && @event.End.DateTime.HasValue)
                    {
                        var tr = new TextRange(string.Format(" - {0}-{1}: {2}\n",
                            @event.Start.DateTime.Value.ToString("HH:mm"),
                            @event.End.DateTime.Value.ToString("HH:mm"),
                            @event.Summary));
                        tr.Format.FontHeight = 16;
                        tr.Format.LatinFont = new TextFont("Lucida Handwriting");
                        tp.TextRanges.Append(tr);
                        
                    }
                }
            }

            return result.ToString();
        }
    }
}