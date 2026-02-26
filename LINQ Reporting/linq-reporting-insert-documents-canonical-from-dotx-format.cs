using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTX template (create + load lifecycle)
        Document template = new Document("Template.dotx");

        // Prepare a data source that will be referenced in the template.
        // The template can use <<[ds.Title]>> and a foreach loop over ds.Items.
        var dataSource = new
        {
            Title = "Quarterly Report",
            Items = GetReportItems()
        };

        // Build the report using the ReportingEngine.
        // The third parameter ("ds") is the name used inside the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "ds");

        // Save the populated document (save lifecycle)
        template.Save("QuarterlyReport.docx");
    }

    // Generates a collection of items using LINQ – this mimics a typical data retrieval scenario.
    private static List<ReportItem> GetReportItems()
    {
        return Enumerable.Range(1, 5)
                         .Select(i => new ReportItem
                         {
                             Name = $"Product {i}",
                             Quantity = i * 10,
                             Price = i * 2.5m
                         })
                         .ToList();
    }

    // Simple POCO that matches the fields used in the template.
    public class ReportItem
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }
}
