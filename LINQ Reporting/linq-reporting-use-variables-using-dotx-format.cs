using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDotxVariableExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOTX template that contains a DOCVARIABLE field like <<[MyVar]>> or a DOCVARIABLE field.
            Document template = new Document(@"Templates\ReportTemplate.dotx");

            // Set a document variable that can be referenced from the template.
            // The variable name is case‑insensitive.
            template.Variables["MyVar"] = "Aspose.Words Reporting";

            // Create a data source object (any non‑dynamic .NET type). Here we use a simple POCO.
            var data = new ReportData
            {
                Title = "Quarterly Sales",
                Date = DateTime.Today,
                Items = new List<SalesItem>
                {
                    new SalesItem { Product = "Laptop", Quantity = 12, Price = 999.99 },
                    new SalesItem { Product = "Smartphone", Quantity = 30, Price = 499.50 }
                }
            };

            // Build the report using the ReportingEngine.
            // The second parameter is the data source; the third parameter is the name used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "ds");

            // Save the populated document in the desired format (DOCX, PDF, etc.).
            template.Save(@"Output\QuarterlyReport.docx", SaveFormat.Docx);
        }
    }

    // Simple POCO used as the data source for the LINQ reporting engine.
    public class ReportData
    {
        public string Title { get; set; }
        public DateTime Date { get; set; }
        public List<SalesItem> Items { get; set; }
    }

    public class SalesItem
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
    }
}
