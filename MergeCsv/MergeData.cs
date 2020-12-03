using System;
using System.Diagnostics;

namespace MergeCsv
{
    class MergeAll
    {
        static void Main()
        {
            Console.WriteLine("Anything special to convert?");
            string whatToConvert = Console.ReadLine();
            Console.Write("Specify CSV dir: ");
            string inputDir = Console.ReadLine();
            Console.WriteLine("Date to convert: ");
            Console.Write("Start date: ");
            string start = Console.ReadLine();
            DateTime startDate = Convert.ToDateTime(start + " 12:00:00 AM");
            Console.Write("End date: ");
            string end = Console.ReadLine();
            DateTime endDate = DateTime.Parse(end + " 11:59:00 PM");
            if (startDate > endDate)
            {
                Console.WriteLine("EndDate > StartDate");
                Environment.Exit(0);
            }
            Console.Write("Specify Output path and file destination: ");
            string outputFile = Console.ReadLine();

            switch (whatToConvert)
            {
                case "a1":
                    var stopwatch = Stopwatch.StartNew();
                    stopwatch.Start();
                    Merge.A1(inputDir, outputFile, startDate, endDate);
                    Merge.Average(outputFile);
                    Console.WriteLine("Chart");
                    Chart.CreateChart(outputFile, outputFile, start, end);
                    stopwatch.Stop();
                    Console.WriteLine("Elapsed Time: " + stopwatch.Elapsed);
                    break;
                case "chart":
                    Chart.CreateChart(inputDir + ".csv", outputFile, start, end);
                    break;
                default:
                    Merge.Any(inputDir, outputFile);
                    break;
            }
        }
    }
}