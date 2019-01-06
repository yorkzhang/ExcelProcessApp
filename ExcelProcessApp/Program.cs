using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ExcelProcessApp
{
    class Program
    {

        static void Main(string[] args)
        {

            Console.WriteLine("Please enter date or date range of the files you want to process(Format:2018.09.14 or 2018.09.14-2018.09.16). If you input nothing and then press enter, it means yesterday. ");

            string inputDate = Console.ReadLine();
            DateTime date = DateTime.MinValue;
            List<DateTime> days = new List<DateTime>();
            if (string.IsNullOrWhiteSpace(inputDate))
            {
                date = DateTime.Today.AddDays(-1);
            }
            else if (inputDate.Contains("-"))
            {
                try
                {
                    string[] dateRangeStrings = inputDate.Split('-');
                    DateTime startDate = DateTime.Parse(dateRangeStrings[0]);
                    DateTime endDate = DateTime.Parse(dateRangeStrings[1]);
                    for (DateTime _date = startDate; _date.Date <= endDate.Date; _date = _date.AddDays(1))
                    {
                        days.Add(_date);
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine($"Wrong date format. If you believe it is a bug, send the below error to author of this program: {e.ToString()}");
                    Environment.Exit(0);
                }
            }
            else
            {
                try
                {
                    date = DateTime.Parse(inputDate);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Wrong date format. If you believe it is a bug, send the below error to author of this program: {e.ToString()}");
                    Environment.Exit(0);
                }
            }
            if (days.Any())
            {
                foreach (var day in days)
                {
                    Console.WriteLine($"*** Now start processing the date for {day.ToShortDateString()}. ***");
                    Console.ReadLine();
                    ProcessData(day);
                }

            }
            else
            {
                Console.WriteLine($"*** Now start processing the date for {date.ToShortDateString()}. ***");
                Console.ReadLine();
                ProcessData(date);
            }
        }

        private static void ProcessData(DateTime date)
        {
            Dictionary<string, bool> fileRecordsDict = new Dictionary<string, bool>()
            {
                { "Data+Summary-Summary(IOS+ANDROID)", false },
                { "Data+Summary-IOS", false },
                { "Data+Summary-ANDROID", false },
                { "新注册用户留存跟踪", false },
                { "流水总况[金额].xls", false },
                { "流水总况[金额] (1).xls", false },
                { "流水总况[金额] (2).xls", false },
                //{ "PVP比赛模式参与度", false },
                //{ "PVE比赛模式参与度", false },

            };
            long accumulatedRegisteredUserCount = 0;
            long dau = 0;
            long newlyRegisteredUserCount = 0;
            decimal averageOnlineDurationPerPerson = 0;

            List<string> unvariableList = new List<string>();
            Console.WriteLine("Now Reading the data of the file contains \"Data+Summary-Summary(IOS+ANDROID)\". Press S Key to skip, other key to continue.");
            ConsoleKey key = Console.ReadKey(true).Key;
            if (key != ConsoleKey.S)
            {
                accumulatedRegisteredUserCount = long.Parse(GetElementValue("Data+Summary-Summary(IOS+ANDROID)", 2, date));
                dau = long.Parse(GetElementValue("Data+Summary-Summary(IOS+ANDROID)", 3, date));
                newlyRegisteredUserCount = long.Parse(GetElementValue("Data+Summary-Summary(IOS+ANDROID)", 4, date));
                averageOnlineDurationPerPerson = decimal.Parse(GetElementValue("Data+Summary-Summary(IOS+ANDROID)", 6, date));
                Console.WriteLine($"Accumulated Registered User: {accumulatedRegisteredUserCount}\nDAU: {dau} \nNewly Registered User: {newlyRegisteredUserCount} \n人均在线时长(分钟): {averageOnlineDurationPerPerson}\n");
                fileRecordsDict["Data+Summary-Summary(IOS+ANDROID)"] = true;
            }
            else
            {
                Console.WriteLine("Skipped reading this file");
            }

            Console.WriteLine("Now Reading the data of the file contains \"Data+Summary-IOS\". Press S Key to skip, other key to continue.");
            long dauIos = 0;
            long newlyRegisteredUserCountIos = 0;
            if (Console.ReadKey(true).Key != ConsoleKey.S)
            {
                newlyRegisteredUserCountIos = long.Parse(GetElementValue("Data+Summary-IOS", 4, date));
                var temp = long.Parse(GetElementValue("Data+Summary-IOS", 4, date));
                dauIos = long.Parse(GetElementValue("Data+Summary-IOS", 3, date));
                Console.WriteLine($"Daily Install - iOS: {newlyRegisteredUserCountIos} \nDAU-iOS: {dauIos}");
                fileRecordsDict["Data+Summary-IOS"] = true;
            }
            else
            {
                Console.WriteLine("Skipped reading this file");
            }

            Console.WriteLine("Now Reading the data of the file contains \"Data+Summary-ANDROID\". Press S Key to skip, other key to continue.");
            long dauAndroid = 0;
            long newlyRegisteredUserCountAndroid = 0;
            if (Console.ReadKey(true).Key != ConsoleKey.S)
            {
                newlyRegisteredUserCountAndroid = long.Parse(GetElementValue("Data+Summary-ANDROID", 4, date));
                dauAndroid = long.Parse(GetElementValue("Data+Summary-ANDROID", 3, date));
                Console.WriteLine($"Daily Install - Android: {newlyRegisteredUserCountAndroid} \nDAU-Android: {dauAndroid}");
                fileRecordsDict["Data+Summary-ANDROID"] = true;
            }
            else
            {
                Console.WriteLine("Skipped reading this file");
            }


            Console.WriteLine("Now Reading the data of the file contains 新注册用户留存跟踪. Press S Key to skip, other key to continue.");
            var retentionPercentValues = new List<string>();
            if (Console.ReadKey(true).Key != ConsoleKey.S)
            {
                retentionPercentValues = GetRetentionPercentValues(date);
                fileRecordsDict["新注册用户留存跟踪"] = true;
            }
            else
            {
                Console.WriteLine("Skipped reading this file");
            }

            Console.WriteLine("Now Reading the data of 流水总况[金额].xls. Press S Key to skip, other key to continue.");
            int playersAllPayedUsersCount = 0;
            long dailyRevRMB = 0;

            if (Console.ReadKey(true).Key != ConsoleKey.S)
            {
                playersAllPayedUsersCount = int.Parse(GetElementValue("流水总况[金额].xls", 2, date));
                dailyRevRMB = long.Parse(GetElementValue("流水总况[金额].xls", 3, date));
                Console.WriteLine($"付费用户数Payers-All: {playersAllPayedUsersCount}\n流水收入总额Daily Rev-Rev-RMB: {dailyRevRMB}");
                fileRecordsDict["流水总况[金额].xls"] = true;
            }
            else
            {
                Console.WriteLine("Skipped reading this file");
            }


            Console.WriteLine("Now Reading the data of 流水总况[金额] (1).xls. Press S Key to skip, other key to continue.");
            int PayersiOSPayedUsersCount = 0;
            long dailyReviOSRMB = 0;
            if (Console.ReadKey(true).Key != ConsoleKey.S)
            {
                PayersiOSPayedUsersCount = int.Parse(GetElementValue("流水总况[金额] (1).xls", 2, date));
                dailyReviOSRMB = long.Parse(GetElementValue("流水总况[金额] (1).xls", 3, date));
                Console.WriteLine($"付费用户数Payers-iOS: {PayersiOSPayedUsersCount}\n流水收入总额Daily Rev-iOS Rev-RMB: {dailyReviOSRMB}");
                fileRecordsDict["流水总况[金额] (1).xls"] = true;
            }
            else
            {
                Console.WriteLine("Skipped reading this file");
            }

            Console.WriteLine("Now Reading the data of 流水总况[金额] (2).xls. Press S Key to skip, other key to continue.");
            long playersAndoridPayedUserCount = 0;
            if (Console.ReadKey(true).Key != ConsoleKey.S)
            {
                playersAndoridPayedUserCount = long.Parse(GetElementValue("流水总况[金额] (2).xls", 2, date));
                Console.WriteLine($" 付费用户数Payers: {playersAndoridPayedUserCount}");
                fileRecordsDict["流水总况[金额] (2).xls"] = true;
            }
            else
            {
                Console.WriteLine("Skipped reading this file");
            }


            //Dictionary<string, string> pvpValues = new Dictionary<string, string>();
            //long vsAttackValue = 0;
            //long pvpValue = 0;
            //long leagueValue = 0;
            //Console.WriteLine("Now Reading the data of the file contains \"PVP比赛模式参与度\", press ENTER to continue, S and Enter to skip");
            //if (!Console.ReadLine().Equals("s", StringComparison.OrdinalIgnoreCase))
            //{
            //    pvpValues = GetMultipleElementsValues("PVP比赛模式参与度", 1, 3, false);
            //    vsAttackValue = long.Parse(pvpValues["VS Attack"]);
            //    pvpValue = long.Parse(pvpValues["PVP"]);
            //    leagueValue = long.Parse(pvpValues["League"]);
            //    Console.WriteLine($"VS Attack:{vsAttackValue} \nPVP:{pvpValue} \nLeague: {leagueValue}");
            //    fileRecordsDict["PVP比赛模式参与度"] = true;
            //}
            //else
            //{
            //    Console.WriteLine("Skipped reading this file");
            //}

            //Console.WriteLine("Now Reading the data of the file contains \"PVE比赛模式参与度\", press ENTER to continue，S and Enter to skip");
            //Dictionary<string, string> pveValues = new Dictionary<string, string>();
            //long allPVEValue = 0;
            //if (!Console.ReadLine().Equals("s", StringComparison.OrdinalIgnoreCase))
            //{
            //    pveValues = GetMultipleElementsValues("PVE比赛模式参与度", 2, 4, true);
            //    allPVEValue = long.Parse(pveValues["&nbsp;"]);
            //    Console.WriteLine($"All PVE: {allPVEValue}");
            //    fileRecordsDict["PVE比赛模式参与度"] = true;
            //}
            //else
            //{
            //    Console.WriteLine("Skipped reading this file");
            //}



            string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\Workbooks";
            string filename = Directory.GetFiles(path, $"*FMC_ChinaPubReport*.*").First();
            var fileInfo = new FileInfo(filename);
            Console.WriteLine($"Will begin writing excel file. Please confirm the filename is {filename}. Enter to continue. Ctrl + C to exit. ");
            Console.ReadLine();

            using (var p = new ExcelPackage(fileInfo))
            {
                //Get the Worksheet created in the previous codesample. 
                Console.WriteLine("Now writing data to  Traffic-D Worksheet. Enter to continue. ");
                Console.ReadLine();
                ExcelWorksheet traffic_d_worksheet = p.Workbook.Worksheets["Traffic-D"];
                ExcelRange dateColumn_traffic_d = traffic_d_worksheet.Cells["C93:C"];
                foreach (var cell in dateColumn_traffic_d)
                {
                    try
                    {
                        var cellDateTime = DateTime.FromOADate(double.Parse(cell.Value.ToString()));
                        if (cellDateTime.Date == date.Date)
                        {
                            if (fileRecordsDict["Data+Summary-Summary(IOS+ANDROID)"])
                            {
                                Console.WriteLine($"Writing ACCUMU INSTALL data:{accumulatedRegisteredUserCount}. ");
                                traffic_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 1].Value = accumulatedRegisteredUserCount;
                                Console.WriteLine($"Writing Daily Install - Total data: {newlyRegisteredUserCount}. ");
                                traffic_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 2].Value = newlyRegisteredUserCount;
                                Console.WriteLine($"Writing DAU:{dau}. ");
                                traffic_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 10].Value = dau;
                                Console.WriteLine($"Writing playingtime:{averageOnlineDurationPerPerson}. ");
                                traffic_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 21].Value = averageOnlineDurationPerPerson;
                            }
                            if (fileRecordsDict["Data+Summary-IOS"])
                            {
                                Console.WriteLine($"Writing DAU-IOS:{dauIos}. ");
                                traffic_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 12].Value = dauIos;
                                Console.WriteLine($"Writing Newly Registered User Count IOS:{newlyRegisteredUserCountIos}. ");
                                traffic_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 4].Value = newlyRegisteredUserCountIos;
                            }

                            if (fileRecordsDict["Data+Summary-ANDROID"])
                            {
                                Console.WriteLine($"Writing Dau Android:{dauAndroid}. ");
                                traffic_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 15].Value = dauAndroid;
                                Console.WriteLine($"Writing Newly Registered User Count Android:{newlyRegisteredUserCountAndroid}.");
                                traffic_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 7].Value = newlyRegisteredUserCountAndroid;

                            }

                            if (fileRecordsDict["新注册用户留存跟踪"])
                            {
                                Console.WriteLine($"Writing 新注册用户留存跟踪:");

                                int i = 0;
                                foreach (var value in retentionPercentValues)
                                {
                                    traffic_d_worksheet.Cells[cell.Start.Row - 1 - i, cell.Start.Column + 22 + i].Value = decimal.Parse(value.TrimEnd(new char[] { '%', ' ' })) / 100M; ;
                                    Console.WriteLine($"D{i + 2} : {value}");
                                    i++;
                                }
                            }

                            try
                            {
                                p.Save();
                                Console.WriteLine("Writing Completed.");
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Did you forget to close the Excel before? Close it and run again!");
                            }

                            break;
                        }
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                }


                Console.WriteLine("Now writing data to  Rev-D Worksheet. Enter to continue. ");
                Console.ReadLine();
                ExcelWorksheet rev_d_worksheet = p.Workbook.Worksheets["Rev-D"];
                ExcelRange dateColumn_rev_d = rev_d_worksheet.Cells["C93:C"];
                foreach (var cell in dateColumn_rev_d)
                {
                    try
                    {
                        var cellDateTime = DateTime.FromOADate(double.Parse(cell.Value.ToString()));
                        if (cellDateTime.Date == date.Date)
                        {
                            if (fileRecordsDict["流水总况[金额].xls"])
                            {
                                Console.WriteLine($"Writing Payers-All data:{playersAllPayedUsersCount}.");
                                rev_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 14].Value = playersAllPayedUsersCount;
                                Console.WriteLine($"Writing Daily Rev-Rev-RMB data:{dailyRevRMB}.");
                                rev_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 1].Value = dailyRevRMB;
                            }

                            if (fileRecordsDict["流水总况[金额] (1).xls"])
                            {
                                Console.WriteLine($"Writing Payers-IOS data:{PayersiOSPayedUsersCount}.");
                                rev_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 15].Value = PayersiOSPayedUsersCount;
                                Console.WriteLine($"Writing Daily Rev-iOS Rev-RMB. data:{dailyReviOSRMB}. ");
                                rev_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 4].Value = dailyReviOSRMB;
                            }
                            if (fileRecordsDict["流水总况[金额] (2).xls"])
                            {
                                Console.WriteLine($"Writing Payers-Android data:{playersAndoridPayedUserCount}.");
                                rev_d_worksheet.Cells[cell.Start.Row, cell.Start.Column + 16].Value = playersAndoridPayedUserCount;
                            }
                            try
                            {
                                p.Save();
                                Console.WriteLine("Writing Completed.");
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Did you forget to close the Excel before? Close it and run again!");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }
                }

                //Console.WriteLine("Now writing data to  Engagement Worksheet. Enter to continue. ");
                //Console.ReadLine();
                //ExcelWorksheet engagement_worksheet = p.Workbook.Worksheets["Engagement"];
                //ExcelRange dateColumn_engagement = engagement_worksheet.Cells["A2:A"];
                //foreach (var cell in dateColumn_engagement)
                //{
                //    try
                //    {
                //        var cellDateTime = DateTime.FromOADate(double.Parse(cell.Value.ToString()));
                //        if (cellDateTime.Date == date.Date)
                //        {
                //            if (fileRecordsDict["PVP比赛模式参与度"])
                //            {
                //                Console.WriteLine($"Writing VS Attack data:{vsAttackValue}. ");
                //                engagement_worksheet.Cells[cell.Start.Row, cell.Start.Column + 2].Value = vsAttackValue;
                //                Console.WriteLine($"Writing PVP data:{pvpValue}. ");
                //                engagement_worksheet.Cells[cell.Start.Row, cell.Start.Column + 3].Value = pvpValue;
                //                Console.WriteLine($"Writing League data:{leagueValue}. ");
                //                engagement_worksheet.Cells[cell.Start.Row, cell.Start.Column + 4].Value = leagueValue;

                //            }
                //            if (fileRecordsDict["PVE比赛模式参与度"])
                //            {
                //                Console.WriteLine($"Writing All PVE data:{allPVEValue}. ");
                //                engagement_worksheet.Cells[cell.Start.Row - 1, cell.Start.Column + 1].Value = allPVEValue;
                //            }
                //            try
                //            {
                //                p.Save();
                //                Console.WriteLine("Writing Completed.");
                //            }
                //            catch (Exception e)
                //            {
                //                Console.WriteLine("Did you forget to close the Excel before? Close it and run again!");
                //            }
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        continue;
                //    }
                //}


            }
        }

        private static HtmlNode GenerateTable(string keyword)
        {
            string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\Workbooks";
            HtmlNode table = null;
            try
            {
                var filename = Directory.GetFiles(path, $"*{keyword}*.*").First();
                Console.WriteLine("Opening file " + path + "...");
                HtmlDocument doc = new HtmlDocument();
                Console.WriteLine("Loading data...");
                doc.Load(filename);
                table = doc.DocumentNode.SelectNodes("//table").First();
            }
            catch (Exception e)
            {
                Console.WriteLine($"Failed to load the file name contains \"{keyword}\" or the data in it. Please check if the file exists or if the data meets the original requirement. System error message{e.ToString()}");
            }
            return table;
        }

        public static string GetElementValue(string keyword, int tableCellIndex, DateTime date)
        {
            HtmlNode table = GenerateTable(keyword);
            var trs = table.SelectNodes("//tr").ToList<HtmlNode>();
            string cellText = string.Empty;
            foreach (var tr in trs)
            {
                if (tr.ChildNodes[0].InnerText == date.ToString("yyyy.MM.dd") || tr.ChildNodes[0].InnerText == date.ToString("yyyy-MM-dd"))
                {
                    cellText = tr.ChildNodes[tableCellIndex].InnerText;
                    break;
                }
            }
            return cellText;
        }
        public static Dictionary<string, string> GetMultipleElementsValues(string keyword, int index1, int index2, bool? isDelayed, DateTime date)
        {
            HtmlNode table = GenerateTable(keyword);
            var trs = table.SelectNodes("//tr").ToList<HtmlNode>();
            Dictionary<string, string> cellsTextValues = new Dictionary<string, string>();
            if (isDelayed.Value)
            {
                var previousDate = date.AddDays(-1);
                foreach (var tr in trs)
                {
                    if (tr.ChildNodes[0].InnerText == previousDate.ToString("yyyy.MM.dd") || tr.ChildNodes[0].InnerText == previousDate.ToString("yyyy-MM-dd"))
                    {
                        cellsTextValues.Add(tr.ChildNodes[index1].InnerText, tr.ChildNodes[index2].InnerText);
                    }
                }
            }

            else
            {
                foreach (var tr in trs)
                {
                    if (tr.ChildNodes[0].InnerText == date.ToString("yyyy.MM.dd") || tr.ChildNodes[0].InnerText == date.ToString("yyyy-MM-dd"))
                    {
                        cellsTextValues.Add(tr.ChildNodes[index1].InnerText, tr.ChildNodes[index2].InnerText);
                    }
                }
            }
            return cellsTextValues;
        }
        public static List<string> GetRetentionPercentValues(DateTime date)
        {
            HtmlNode table = GenerateTable("新注册用户留存跟踪");


            List<HtmlNode> trs = table.SelectNodes("//tr").ToList<HtmlNode>();
            var retentionPercentDate = date.AddDays(-1);
            HtmlNode tdForDateAndPercent = trs[0].SelectNodes("td").Where(x => x.InnerText.Contains(retentionPercentDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)) && x.InnerText.Contains(date.ToString("留存率", CultureInfo.InvariantCulture))).First();
            int tdIndex = trs[0].SelectNodes("td").GetNodeIndex(tdForDateAndPercent);
            List<string> retentionPercentValues = new List<string>();

            for (int i = 0; i <= 28; i++)
            {
                string retentionPercentValue = trs[i + 2].ChildNodes[tdIndex - 2 * i].InnerText;
                retentionPercentValues.Add(
                    retentionPercentValue
                    );
                Console.WriteLine($"D{i + 2} : {retentionPercentValue}");
            }

            return retentionPercentValues;
        }
    }
}

