using System.IO;
using Microsoft.Office.Interop.Excel;

namespace MergeCsv
{
    public class Chart
    {
        static string chartStartCell = "E1";
        static string chartEndCell = "F15";
        static string graphTitle = "KPI Averages ";
        static string xAxis = "Tests";
        static string yAxis = "Time";
        //static string _columnLetter;

        public static void CreateChart(string inputFolder, string outputFile, string start, string end)
        {
            if (File.Exists(outputFile + @"Chart.xlsx"))
            {
                File.Delete(outputFile + @"Chart.xlsx");
            }
            // Open Excel and get first worksheet.
            var application = new Application();
            var workbook = application.Workbooks.Open(inputFolder);
            var worksheet = workbook.Worksheets[1] as Worksheet;
            Range xlRange = worksheet.UsedRange;
            var rowCount = xlRange.Rows.Count.ToString();
            var columnCount = xlRange.Columns.Count;
            // switch (columnCount)
            // {
            //     case 1:
            //         _columnLetter = "A";
            //         break;
            //     case 2:
            //         _columnLetter = "B";
            //         break;
            //     case 3:
            //         _columnLetter = "C";
            //         break;
            //     case 4:
            //         _columnLetter = "D";
            //         break;
            //     case 5:
            //         _columnLetter = "E";
            //         break;
            //     case 6:
            //         _columnLetter = "F";
            //         break;
            //     case 7:
            //         _columnLetter = "G";
            //         break;
            //     default:
            //         Console.WriteLine("Too much data.");
            //         _columnLetter = "A";
            //         break;
            // }

            // Add chart.
            var charts = worksheet.ChartObjects() as ChartObjects;
            var chartObject = charts.Add(60, 10, 800, 400);
            var chart = chartObject.Chart;

            // Set chart range.
            //chartEndCell = _columnLetter + rowCount;
            var range = worksheet.Range[chartStartCell, chartEndCell];
            chart.SetSourceData(range);

            // Set chart properties.
            graphTitle = graphTitle + $"{start} - {end}";
            var seriesCollection = (SeriesCollection)chart.SeriesCollection();
            Series s1 = seriesCollection.NewSeries();
            s1.Name = "Border value";
            s1.MarkerStyle = XlMarkerStyle.xlMarkerStyleAutomatic;
            s1.Values = worksheet.Range["G1", "G15"];
            chart.ChartType = XlChartType.xlBarClustered;
            chart.ApplyDataLabels();
            chart.ChartWizard(
                Title: graphTitle,
                CategoryTitle: xAxis,
                ValueTitle: yAxis);

            // Save.
            workbook.SaveAs(outputFile + @"Chart.xlsx");
        }
    }
}