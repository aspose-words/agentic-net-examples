// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingPrint
{
    // Simple data source class for the report
    public class ReportData
    {
        public string CustomerName { get; set; }
        public DateTime OrderDate { get; set; }
        public decimal TotalAmount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCM template that contains LINQ reporting tags
            const string templatePath = @"C:\Templates\ReportTemplate.docm";

            // Load the DOCM document (lifecycle: load)
            Document doc = new Document(templatePath);

            // Prepare the data source for the report
            ReportData data = new ReportData
            {
                CustomerName = "John Doe",
                OrderDate = DateTime.Today,
                TotalAmount = 1234.56m
            };

            // Create the reporting engine (LINQ reporting)
            ReportingEngine engine = new ReportingEngine();

            // Populate the template with data (reporting engine builds the report)
            // The third argument is the name used inside the template to reference the data source.
            engine.BuildReport(doc, data, "ds");

            // Optional: configure printer settings (e.g., print only first two pages)
            PrinterSettings printerSettings = new PrinterSettings
            {
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 2
            };

            // Print the populated document using the specified printer settings
            doc.Print(printerSettings, "LINQ Report Document");

            // If you need to save the populated document before printing, use the save lifecycle method:
            // doc.Save(@"C:\Output\PopulatedReport.docx");
        }
    }
}
