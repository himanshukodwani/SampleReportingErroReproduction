using System;
using System.IO;

namespace SampleReporting
{
    //SharpLightReporting is a reporting library which is based on SpreadsheetLight which is an excel library which is based on DocumentFormat.OpenXML
    internal class Program
    {
        private static string TemplateFilePath = System.IO.Directory.GetCurrentDirectory() + System.IO.Path.DirectorySeparatorChar.ToString() + "Template" + System.IO.Path.DirectorySeparatorChar + "InvoiceTemplate.xlsx";
        private static string OutputFilePath = System.IO.Directory.GetCurrentDirectory() + System.IO.Path.DirectorySeparatorChar.ToString() + "Template" + System.IO.Path.DirectorySeparatorChar + "InvoiceGenerated.xlsx";

        private static void Main(string[] args)
        {
            Console.WriteLine("Welcome to Excel Invoice Generation Sample");
            Console.WriteLine(TemplateFilePath);
            if (File.Exists(TemplateFilePath))
            {
                InvoiceReportModel model = new InvoiceReportModel(); //Contains all the data required to fill the invoice template

                SharpLightReporting.ReportEngine reportEngine = new SharpLightReporting.ReportEngine();
                reportEngine.ProcessReport(TemplateFilePath, OutputFilePath, model);
            }
            else
            {
                Console.WriteLine("File not found");
            }
            Console.WriteLine("Done");
            Console.ReadLine();
        }
    }
}