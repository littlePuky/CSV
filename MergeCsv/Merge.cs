using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
        public static string OutputFile;

        public static void A1(string inputFolder, string outputFile, DateTime start, DateTime end)
        {
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }

            FolderPath = inputFolder;
            List<string> STB5019 = new List<string>();
            STB5019.Add("STB,FileName,Date,Time");
            List<string> STB5020 = new List<string>();
            //STB5020.Add("STB,FileName,Date,Time");
            string[] logs = Directory.GetFiles(FolderPath, "TestRun.log", SearchOption.AllDirectories);

            foreach (var log in logs)
            {
                var fileDate = File.GetLastWriteTime(log);
                if (!(fileDate >= start && fileDate <= end))
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

        public static void Average(string input, string output)
        {
            if (File.Exists(output + ".xlsx"))
            {
                File.Delete(output + ".xlsx");
            }

            Console.WriteLine("Average");
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(input);
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;
            x.Cells[1, 5] = "KPI_1_Average";
            x.Cells[1, 7] = "0.5";
            x.Range["F1"].Formula = "=AVERAGEIF(B:BT,\"KPI_1_PageTransition_results\",D:D)";
            x.Cells[2, 5] = "KPI_2_ZapTime";
            x.Cells[2, 7] = "3.5";
            x.Range["F2"].Formula = "=AVERAGEIF(B:BT,\"KPI_2_ZapTime_results\",D:D)";
            x.Cells[3, 5] = "KPI_3_LinearEPGDetailPage";
            x.Cells[3, 7] = "1.0";
            x.Range["F3"].Formula = "=AVERAGEIF(B:BT,\"KPI_3_LinearEPGDetailPage_results\",D:D)";
            x.Cells[4, 5] = "KPI_8_MenuDisplayTime";
            x.Cells[4, 7] = "0.6";
            x.Range["F4"].Formula = "=AVERAGEIF(B:BT,\"KPI_8_MenuDisplayTime_results\",D:D)";
            x.Cells[5, 5] = "KPI_14_Average";
            x.Cells[5, 7] = "2.0";
            x.Range["F5"].Formula = "=AVERAGEIF(B:BT,\"KPI_14_AverageEPGNavigation_results\",D:D)";
            x.Cells[6, 5] = "KPI_15_Average";
            x.Cells[6, 7] = "1.5";
            x.Range["F6"].Formula = "=AVERAGEIF(B:BT,\"KPI_15_AverageEPGPageChange_results\",D:D)";
            x.Cells[7, 5] = "KPI_16_Average";
            x.Cells[7, 7] = "1.0";
            x.Range["F7"].Formula = "=AVERAGEIF(B:BT,\"KPI_16_AverageEPGMenu_results\",D:D)";
            x.Cells[8, 5] = "KPI_19_AverageSearchTime";
            x.Cells[8, 7] = "0.1";
            x.Range["F8"].Formula = "=AVERAGEIF(B:BT,\"KPI_19_AverageSearchTime_results\",D:D)";
            x.Cells[9, 5] = "KPI_20_AverageSearchNavigation";
            x.Cells[9, 7] = "1.0";
            x.Range["F9"].Formula = "=AVERAGEIF(B:BT,\"KPI_20_AverageSearchNavigation_results\",D:D)";
            x.Cells[10, 5] = "KPI_25_ColdBoot_UntilVisualFeedback";
            x.Cells[10, 7] = "70.0";
            x.Range["F10"].Formula = "=AVERAGEIF(B:BT,\"KPI_25_ColdBoot_UntilVisualFeedback_results\",D:D)";
            x.Cells[11, 5] = "KPI_26_ColdBoot_UntilStream";
            x.Cells[11, 7] = "85.0";
            x.Range["F11"].Formula = "=AVERAGEIF(B:BT,\"KPI_26_ColdBoot_UntilStream_results\",D:D)";
            x.Cells[12, 5] = "KPI_27_ColdBoot_UntilHomepage";
            x.Cells[12, 7] = "80.0";
            x.Range["F12"].Formula = "=AVERAGEIF(B:BT,\"KPI_27_ColdBoot_UntilHomepage_results\",D:D)";
            x.Cells[13, 5] = "KPI_28_Average";
            x.Cells[13, 7] = "2.0";
            x.Range["F13"].Formula = "=AVERAGEIF(B:BT,\"KPI_28_Standby_UntilVisualFeedback_results\",D:D)";
            x.Cells[14, 5] = "KPI_29_Standby";
            x.Cells[14, 7] = "10.0";
            x.Range["F14"].Formula = "=AVERAGEIF(B:BT,\"KPI_29_Standby_UntilHome_results\",D:D)";
            x.Cells[15, 5] = "KPI_30_Standby";
            x.Cells[15, 7] = "15.0";
            x.Range["F15"].Formula = "=AVERAGEIF(B:BT,\"KPI_30_Standby_UntilStream_results\",D:D)";
            x.Cells[16, 5] = "ZAP_Video";
            x.Cells[16, 7] = "1.2";
            x.Range["F16"].Formula = "=AVERAGEIF(B:BT,\"KPI_2_ZapTimeNew_video_results\",D:D)";
            x.Cells[17, 5] = "ZAP_Audio";
            x.Cells[17, 7] = "3.0";
            x.Range["F17"].Formula = "=AVERAGEIF(B:BT,\"KPI_2_ZapTimeNew_audio_results\",D:D)";
            sheet.SaveAs(output + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Local: Type.Missing);
            File.Delete(input);
            excel.Quit();
        }

        public static void Any(string input, string output)
        {
            OutputFile = output + ".csv";
            var csvWriter = new StreamWriter(OutputFile);
            string[] files = Directory.GetFiles(input, "*.csv", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                string[] allLines = File.ReadAllLines(file);
                foreach (var line in allLines)
                {
                    var info = new FileInfo(file);
                    Modified = info.LastWriteTime.ToString();
                    var data = line;
                    FileName = Path.GetFileNameWithoutExtension(file);
                    csvWriter.WriteLine($"{FileName}, {Modified}, {data}");
                    csvWriter.Flush();
                }
            }
        }
    }
}