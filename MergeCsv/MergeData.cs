using System;

namespace MergeCsv
{
    class MergeAll
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Anything special to convert?");
            string whatToConvert = Console.ReadLine();
            Console.Write("Specify CSV dir: ");
            string imputDir = Console.ReadLine();
            Console.WriteLine("Date to convert: ");
            Console.Write("Start date: ");
            string start = Console.ReadLine();
            DateTime startDate = Convert.ToDateTime(start + " 12:00:00 AM");
            Console.Write("End date: ");
            string end = Console.ReadLine();
            DateTime endDate = DateTime.Parse(end + " 11:59:00 PM");
            Console.Write("Specify Output path and file to save: ");
            string outputFile = Console.ReadLine();

            switch (whatToConvert)
            {
                case "a1":
                    Merge.A1(imputDir, outputFile, startDate, endDate);
                    Merge.average(outputFile);
                    Console.WriteLine("Chart");
                    Chart.CreateChart(outputFile, outputFile, start, end);
                    break;
                case "chart":
                    Chart.CreateChart(imputDir + ".csv", outputFile, start, end);
                    break;
                default:
                    Merge.Any();
                    break;
            }
        }
    }
}