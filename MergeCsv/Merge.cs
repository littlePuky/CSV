using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace MergeCsv
{
    public class Merge
    {
        public static string Modified = "";
        public static double AverageData;
        public static string CreationTime;
        public static string FileName = "";
        public static string FolderPath;

        public static void A1(string imputFolder, string outputFile, DateTime start, DateTime end)
        {
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }
            FolderPath = imputFolder;
            List<string> STB5019 = new List<string>();
            STB5019.Add("STB,FileName,Date,Time");
            List<string> STB5020 = new List<string>();
            //STB5020.Add("STB,FileName,Date,Time");
            string[] logs = Directory.GetFiles(FolderPath, "TestRun.log", SearchOption.AllDirectories);

            foreach (var log in logs)
            {
                var fileDate = File.GetLastWriteTime(log);
                if (!(fileDate>=start && fileDate<=end))
                {
                    continue;
                }
                var filePath = Path.GetDirectoryName(log);

                string GetLine(string file1, int line1 = 6)
                {
                    using (var sr = new StreamReader(file1))
                    {
                        for (int i = 1; i < line1; i++)
                            sr.ReadLine();
                        if (sr.ReadLine() == null)
                        {
                            return "null";
                        }

                        return sr.ReadLine();
                    }
                }

                string[] files = Directory.GetFiles(filePath, "*.csv", SearchOption.AllDirectories);
                foreach (var file in files)
                {
                    
                    string[] allLines = File.ReadAllLines(file);
                    FileName = Path.GetFileNameWithoutExtension(file);
                    foreach (var line in allLines)
                    {   
                        FileInfo info = new FileInfo(file);
                        Modified = info.LastWriteTime.ToShortDateString();
                        var asd = File.GetLastWriteTimeUtc(file);
                        var data = double.Parse(line);
                        if (data < 0)
                        {
                            continue;
                        }
                        CreationTime = File.GetCreationTime(file).ToString();
                        
                        if (GetLine(log).Contains("STB_Type5019"))
                        {
                            STB5019.Add($"STB_Type5019,{FileName},{Modified},{data}");
                        }

                        if (GetLine(log).Contains("STB_Type5020"))
                        {
                            STB5020.Add($"STB_Type5020,{FileName},{Modified},{data}");
                        }
                    }
                }
            }

            OutputFile = FolderPath + ".csv";
            var csvWriter = new StreamWriter(outputFile);
            foreach (var line in STB5019)
            {
                csvWriter.WriteLine(line);
                csvWriter.Flush();
            }

            //OutputFile = FolderPath + "5020" + ".csv";
            //var csvWriter1 = new StreamWriter(OutputFile);
            foreach (var line in STB5020)
            {
                csvWriter.WriteLine(line);
                csvWriter.Flush();
            }
            csvWriter.Close();
        }

        public static string OutputFile;

        public static void average(string InputFolder)
        {
            Console.WriteLine("Average");
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(InputFolder);
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;
            x.Cells[1, 5] = "KPI_1_Average";
            x.Range["F1"].Formula = "=AVERAGEIF(B:BT,\"KPI_1_PageTransition_results\",D:D)";
            x.Cells[2, 5] = "KPI_2_ZapTime";
            x.Range["F2"].Formula = "=AVERAGEIF(B:BT,\"KPI_2_ZapTime_results\",D:D)";
            x.Cells[3, 5] = "KPI_3_LinearEPGDetailPage";
            x.Range["F3"].Formula = "=AVERAGEIF(B:BT,\"KPI_3_LinearEPGDetailPage_results\",D:D)";
            x.Cells[4, 5] = "KPI_8_MenuDisplayTime";
            x.Range["F4"].Formula = "=AVERAGEIF(B:BT,\"KPI_8_MenuDisplayTime_results\",D:D)";
            x.Cells[5, 5] = "KPI_14_Average";
            x.Range["F5"].Formula = "=AVERAGEIF(B:BT,\"KPI_14_AverageEPGNavigation_results\",D:D)";
            x.Cells[6, 5] = "KPI_15_Average";
            x.Range["F6"].Formula = "=AVERAGEIF(B:BT,\"KPI_15_AverageEPGPageChange_results\",D:D)";
            x.Cells[7, 5] = "KPI_16_Average";
            x.Range["F7"].Formula = "=AVERAGEIF(B:BT,\"KPI_16_AverageEPGMenu_results\",D:D)";
            x.Cells[8, 5] = "KPI_19_AverageSearchTime";
            x.Range["F8"].Formula = "=AVERAGEIF(B:BT,\"KPI_19_AverageSearchTime_results\",D:D)";
            x.Cells[9, 5] = "KPI_20_AverageSearchNavigation";
            x.Range["F9"].Formula = "=AVERAGEIF(B:BT,\"KPI_20_AverageSearchNavigation_results\",D:D)";
            x.Cells[10, 5] = "KPI_25_ColdBoot_UntilVisualFeedback";
            x.Range["F10"].Formula = "=AVERAGEIF(B:BT,\"KPI_25_ColdBoot_UntilVisualFeedback_results\",D:D)";
            x.Cells[11, 5] = "KPI_26_ColdBoot_UntilStream";
            x.Range["F11"].Formula = "=AVERAGEIF(B:BT,\"KPI_26_ColdBoot_UntilStream_results\",D:D)";
            x.Cells[12, 5] = "KPI_27_ColdBoot_UntilHomepage";
            x.Range["F12"].Formula = "=AVERAGEIF(B:BT,\"KPI_27_ColdBoot_UntilHomepage_results\",D:D)";
            x.Cells[13, 5] = "KPI_28_Average";
            x.Range["F13"].Formula = "=AVERAGEIF(B:BT,\"KPI_28_Standby_UntilVisualFeedback_results\",D:D)";
            x.Cells[14, 5] = "KPI_29_Standby";
            x.Range["F14"].Formula = "=AVERAGEIF(B:BT,\"KPI_29_Standby_UntilHome_results\",D:D)";
            x.Cells[15, 5] = "KPI_30_Standby";
            x.Range["F15"].Formula = "=AVERAGEIF(B:BT,\"KPI_30_Standby_UntilStream_results\",D:D)";
            sheet.Save();
            excel.Quit();
        }
        public static void Any()
        {
            Console.Write("Specify CSV dir: ");
            FolderPath = Console.ReadLine();
            OutputFile = FolderPath + ".csv";
            var csvWriter = new StreamWriter(OutputFile);
            string[] files = Directory.GetFiles(FolderPath, "*.csv", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                string[] allLines = File.ReadAllLines(file);
                foreach (var t in allLines)
                {
                    var info = new FileInfo(file);
                    Modified = info.LastWriteTime.ToString();
                    var data = t;
                    CreationTime = File.GetCreationTime(file).ToString();
                    FileName = Path.GetFileNameWithoutExtension(file);
                    csvWriter.WriteLine($"{FileName}, {Modified}, {data}");
                    csvWriter.Flush();
                }
            }
        }
        
    }
}