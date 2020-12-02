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
            Console.Write("Specify Output path and file to save: ");
            string outputFile = Console.ReadLine();
            switch (whatToConvert)
            {
                case "a1":
                    Merge.A1(imputDir, outputFile);
                    Merge.average(outputFile);
                    Console.WriteLine("Chart");
                    Chart.CreateChart(outputFile, outputFile);
                    break;
                case "chart":
                    Chart.CreateChart(imputDir + ".csv", outputFile);
                    break;
                default:
                    Merge.Any();
                    break;
            }
        }
    }
}