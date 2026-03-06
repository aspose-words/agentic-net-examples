// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTX template from file system.
        Document document = new Document("Template.dotx");

        // Example data source for LINQ reporting.
        var reportData = new
        {
            Title = "Sales Report",
            Date = DateTime.Today,
            Items = new[]
            {
                new { Product = "Apple",  Quantity = 120, Price = 0.5 },
                new { Product = "Banana", Quantity = 85,  Price = 0.3 },
                new { Product = "Cherry", Quantity = 60,  Price = 1.2 }
            }
        };

        // Populate the template with data using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The third argument ("ds") is the name used in the template to reference the data source.
        engine.BuildReport(document, reportData, "ds");

        // Print the resulting document to the default printer.
        document.Print();

        // Optional: print to a specific printer with a page range.
        // PrinterSettings settings = new PrinterSettings
        // {
        //     PrinterName = "Your Printer Name",
        //     PrintRange = PrintRange.SomePages,
        //     FromPage = 1,
        //     ToPage = document.PageCount
        // };
        // document.Print(settings);
    }
}
