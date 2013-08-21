using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Serialization;
using System.IO;
using OfficeOpenXml;

namespace WebPlaylistGrabber
{
    /// <summary>
    /// Custom comparer for getting the equal objects
    /// </summary>
    class DistinctSongEntryComparer : IEqualityComparer<SongEntry>
    {

        public bool Equals(SongEntry x, SongEntry y)
        {
            return x.AirTime == y.AirTime &&
                x.Artist == y.Artist &&
                x.Title == y.Title;
        }

        public int GetHashCode(SongEntry obj)
        {
            return obj.AirTime.GetHashCode() ^
                obj.Artist.GetHashCode() ^
                obj.Title.GetHashCode();
        }
    }

    /// <summary>
    /// A single song-entry
    /// </summary>
    [Serializable]
    public class SongEntry
    {
        /// <summary>
        /// The Artist of the song
        /// </summary>
        public string Artist { get; set; }
        /// <summary>
        /// The time and date, when this song got played
        /// </summary>
        public DateTime AirTime { get; set; }
        /// <summary>
        /// The title of the song
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// A default constructor, so that the class can be serialized
        /// </summary>
        public SongEntry()
        {

        }
    }

    
    //Pretty bad comments, more like a proof of concept!!!
    class Program
    {
        public static List<SongEntry> songs;

        static void Main(string[] args)
        {
            
            songs = new List<SongEntry>();

            //Hardcoded dates for the last week
            List<string> dates = new List<string>();
            dates.Add("12.08.2013");
            dates.Add("13.08.2013");
            dates.Add("14.08.2013");
            dates.Add("15.08.2013");
            dates.Add("16.08.2013");
            dates.Add("17.08.2013");

            int hour_from = 06;
            int hour_to = 22;

            int current_hour = hour_from;
            int current_minute = 00;

            List<SongEntry> finalSongList = new List<SongEntry>();

            foreach (string date in dates)
            {
                songs = new List<SongEntry>();
                current_hour = hour_from;
                current_minute = 00;
                Console.WriteLine("Date: " + date);
                while (current_hour < hour_to)
                {
                    Console.WriteLine(String.Format("Lese {0}:{1}", current_hour, current_minute));

                    string queryURL = String.Format(@"http://www.radiobremen.de/bremenvier/musik/titelsuche/?wrapurl=%2Fbremenvier%2Fmusik%2Ftitelsuche%2F&titelsuche_datum={0}&titelsuche_stunden={1}&titelsuche_minuten={2}&submit_titelsuche=Suchen", date, current_hour, current_minute);
                    var webGet = new HtmlWeb();
                    HtmlDocument webPage = webGet.Load(queryURL, "GET");

                    for (int n = 2; n < 9; n++)
                    {
                        HtmlNode data = webPage.DocumentNode.SelectNodes(String.Format("//*[@id=\"verlauf_inner_content\"]/center/table/tbody/tr[{0}]/td[1]", n)).SingleOrDefault();
                        string time = data.InnerText;
                        data = webPage.DocumentNode.SelectNodes(String.Format("//*[@id=\"verlauf_inner_content\"]/center/table/tbody/tr[{0}]/td[2]", n)).SingleOrDefault();
                        string artist = data.InnerText;
                        data = webPage.DocumentNode.SelectNodes(String.Format("//*[@id=\"verlauf_inner_content\"]/center/table/tbody/tr[{0}]/td[3]", n)).SingleOrDefault();
                        string song = data.InnerText;

                        songs.Add(new SongEntry
                        {
                            AirTime = DateTime.Parse(date + " " + time),
                            Artist = artist,
                            Title = song
                        });
                    }



                    if (current_minute + 20 > 59)
                    {
                        current_hour = current_hour + 1;
                        current_minute = ((current_minute + 20) - 60);
                    }
                    else
                    {
                        current_minute = current_minute + 20;
                    }
                }
                finalSongList.AddRange(songs.Distinct(new DistinctSongEntryComparer()).ToList());

            }

            //Group Songs by Artist and Title.
            var airtimes = (from p in finalSongList
                            group p by new
                            {
                                p.Artist,
                                p.Title
                            } into g
                            select new { 
                                Artist = g.Key.Artist, 
                                Title = g.Key.Title, 
                                dates = from n in g 
                                        select new { 
                                            Time = n.AirTime 
                                        }, 
                                amount = g.Count() }
                            ).OrderBy(o => o.amount);
            

                FileInfo newFile = new FileInfo(@"D:\sample1.xlsx");

                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    // add a new worksheet to the empty workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Songs @ Bremen 4");
                    //Add the headers
                    worksheet.Cells[1, 1].Value = "Künstler";
                    worksheet.Cells[1, 2].Value = "Titel";
                    worksheet.Cells[1, 3].Value = "Anzahl";

                    int cnt = 2;
                    foreach (var airtime in airtimes)
                    {
                        //Add some items...
                        worksheet.Cells[cnt, 1].Value = airtime.Artist;
                        worksheet.Cells[cnt, 2].Value = airtime.Title;
                        worksheet.Cells[cnt, 3].Value = airtime.amount;
                        cnt++;
                    }

                    package.Save();
            }
        }
    }
}
