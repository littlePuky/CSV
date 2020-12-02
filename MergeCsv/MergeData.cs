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
            string dir = Console.ReadLine();
            switch (whatToConvert)
            {
                case "a1":
                    Merge.A1(dir);
                    Merge.average(dir+".csv");
                    Console.WriteLine("Chart");
                    Chart.CreateChart(dir+".csv", dir);
                    break;
                case "chart":
                    Chart.CreateChart(dir+".csv", dir);
                    break;
                default:
                    Merge.Any();
                    break;
            }
        }
    }
}