using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace RealYoutubeApiApp
{
    class Program
    {
        static void Main(string[] args)
        {   
            Console.Title = ("API App");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Clear();

            Console.WriteLine("Search By Channel URL or Video URL -- video || channel");
            string theSearch = Console.ReadLine();

            if (theSearch == "video")
            {
                Console.WriteLine("Input number of URL's!");
                int n = int.Parse(Console.ReadLine());

                string[] arrId = new string[n];
                string[] arrURL = new string[n];
                string[] arrName = new string[n];
                string[] fullURLS = new string[n];

                string[] titles = new string[n];
                string[] uploadDates = new string[n];
                string[] channels = new string[n];
                string[] channelID = new string[n];

                Console.WriteLine("Input URL's!");
                for (int i = 0; i < n; i++)
                {
                    string url = Console.ReadLine();
                    fullURLS[i] = url;
                    string urlToUse = string.Empty;
                    string secLast = string.Empty;
                    if (url.Contains("www.youtube.com"))
                    {

                        string[] urlArr = url.Split("watch?v=").ToArray();
                        urlToUse = urlArr.Last();
                        if (urlToUse.Contains('/'))
                        {
                            urlToUse = urlToUse.Remove(urlToUse.Length - 1);
                        }
                    }
                    arrURL[i] = urlToUse;
                }

                for (int i = 0; i < arrURL.Length; i++)
                {
                    string title = default;
                    string uploadTime = default;
                    string channel = default;
                    string channelId = default;
                    string dateToUse = default;

                    if (i != 0 && arrURL[i] == arrURL[i - 1])
                    {
                        titles[i] = titles[i - 1];
                        uploadDates[i] = uploadDates[i - 1];
                        channels[i] = channels[i - 1];
                        channelID[i] = channelID[i - 1];
                    }
                    else
                    {
                        HttpWebRequest titlerequest = (HttpWebRequest)WebRequest.Create($"https://www.googleapis.com/youtube/v3/videos?part=snippet&id={arrURL[i]}&key=************************************");
                        //first API Key: ************************************
                        //stefan key: ************************************
                        //new Key: ************************************
                        //another one: ************************************
                        // last:  ************************************
                        HttpWebResponse titleresponse = (HttpWebResponse)titlerequest.GetResponse();
                        Stream titlestream = titleresponse.GetResponseStream();
                        StreamReader titlereader = new StreamReader(titlestream);
                        string titlejson = titlereader.ReadToEnd();
                        Newtonsoft.Json.Linq.JObject jObject = Newtonsoft.Json.Linq.JObject.Parse(titlejson);


                        try
                        {
                            title = (string)jObject["items"][0]["snippet"]["title"];
                            uploadTime = (string)jObject["items"][0]["snippet"]["publishedAt"];
                            channel = (string)jObject["items"][0]["snippet"]["channelTitle"];
                            channelId = (string)jObject["items"][0]["snippet"]["channelId"];
                            //Abrufe dazu

                            string[] justDate = uploadTime.Split();
                            string[] dates = justDate[0].Split("/");
                            dateToUse = $"{dates[1]}.{dates[0]}.{dates[2]}";
                        }
                        
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            title = "UNAVAILABLE";
                            uploadTime = "UNAVAILABLE";
                            channel = "UNAVAILABLE";
                            dateToUse = "UNAVAILABLE";
                            channelId = "UNAVAILABLE";
                        }

                        titles[i] = title;
                        uploadDates[i] = dateToUse;
                        channels[i] = channel;
                        channelID[i] = channelId;
                    }
                }
                Console.ForegroundColor = ConsoleColor.Yellow;

                /*  string titlesPath = Path.Combine("..", "..", "..", "titles.txt");
                  string datesPath = Path.Combine("..", "..", "..", "dates.txt");
                  string channelsPath = Path.Combine("..", "..", "..", "channels.txt");*/

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();


                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add();
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "Titles";
                xlWorkSheet.Cells[1, 2] = "Dates";
                xlWorkSheet.Cells[1, 3] = "Channels";
                xlWorkSheet.Cells[1, 4] = "ChannelID";

                Console.WriteLine("Loading... ");
                for (int i = 0; i < titles.Length; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1] = titles[i];
                    xlWorkSheet.Cells[i + 2, 2] = uploadDates[i];
                    xlWorkSheet.Cells[i + 2, 3] = channels[i];
                    xlWorkSheet.Cells[i + 2, 4] = channelID[i];
                    //File.AppendAllText(titlesPath, titles[i].TrimEnd() + Environment.NewLine);
                    //Console.WriteLine(titles[i].TrimEnd());
                }

                //Console.WriteLine("Upload Dates Are: ");
                //for (int i = 0; i < uploadDates.Length; i++)
                {
                    
                    //File.AppendAllText(datesPath, uploadDates[i].TrimEnd() + Environment.NewLine);
                    //Console.WriteLine(uploadDates[i]);
                }

               // Console.WriteLine("Channels Are: ");
                //for (int i = 0; i < channels.Length; i++)
                {
                    
                    //File.AppendAllText(channelsPath, channels[i].TrimEnd() + Environment.NewLine);
                    //Console.WriteLine(channels[i]);
                }

                xlWorkBook.SaveAs(@"D:\From Old Laptop - Work, Uni, Code\SoftUni\AppsForTuneSat\RealYoutubeApiApp\Api.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
            }
            else
            {
                Console.WriteLine("Input number of URL's!");
                int n = int.Parse(Console.ReadLine());

                string[] arrId = new string[n];
                string[] arrURL = new string[n];
                string[] arrName = new string[n];
                string[] fullURLS = new string[n];

                string[] channels = new string[n];
                string[] countries = new string[n];

                Console.WriteLine("Input URL's!");
                for (int i = 0; i < n; i++)
                {
                    string url = Console.ReadLine();
                    fullURLS[i] = url;
                    string urlToUse = string.Empty;
                    string secLast = string.Empty;
                    if (url.Contains("www.youtube.com"))
                    {

                        string[] urlArr = url.Split("/channel/").ToArray();
                        urlToUse = urlArr.Last();
                        if (urlToUse.Contains('/'))
                        {
                            urlToUse = urlToUse.Remove(urlToUse.Length - 1);
                        }
                    }
                    else
                    {
                        urlToUse = url;
                    }
                    arrURL[i] = urlToUse;
                }

                for (int i = 0; i < arrURL.Length; i++)
                {
                    string channel = default;
                    string country = default;

                    if (i != 0 && arrURL[i] == arrURL[i - 1])
                    {
                        channels[i] = channels[i - 1];
                        countries[i] = countries[i - 1];
                    }
                    else
                    {
                        if (arrURL[i] == "UNAVAILABLE")
                        {
                            channels[i] = "UNAVAILABLE";
                            countries[i] = "UNAVAILABLE";
                            continue;
                        }
                        HttpWebRequest titlerequest = (HttpWebRequest)WebRequest.Create($"https://www.googleapis.com/youtube/v3/channels?id={arrURL[i]}&part=snippet&key=************************************");
                        //first API Key: ************************************
                        //second API Key: ************************************
                        HttpWebResponse titleresponse = (HttpWebResponse)titlerequest.GetResponse();
                        Stream titlestream = titleresponse.GetResponseStream();
                        StreamReader titlereader = new StreamReader(titlestream);
                        string titlejson = titlereader.ReadToEnd();
                        Newtonsoft.Json.Linq.JObject jObject = Newtonsoft.Json.Linq.JObject.Parse(titlejson);


                        try
                        {
                            channel = (string)jObject["items"][0]["snippet"]["title"];
                            country = (string)jObject["items"][0]["snippet"]["country"];
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            channel = "UNAVAILABLE";
                            country = "UNAVAILABLE";
                        }

                        channels[i] = channel;
                        countries[i] = country;
                    }
                }
                Console.ForegroundColor = ConsoleColor.Yellow;

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();


                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add();
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "Channels";
                xlWorkSheet.Cells[1, 2] = "URLS";
                xlWorkSheet.Cells[1, 3] = "Country";


                Console.WriteLine("Channels Are: ");
                for (int i = 0; i < channels.Length; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1] = channels[i];
                    xlWorkSheet.Cells[i + 2, 2] = arrURL[i];
                    xlWorkSheet.Cells[i + 2, 3] = countries[i];
                    //File.AppendAllText(channelsPath, channels[i].TrimEnd() + Environment.NewLine);
                    //Console.WriteLine(channels[i]);
                }

                xlWorkBook.SaveAs(@"D:\From Old Laptop - Work, Uni, Code\SoftUni\AppsForTuneSat\RealYoutubeApiApp\Counrtry.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
            }
        }
    }
}
