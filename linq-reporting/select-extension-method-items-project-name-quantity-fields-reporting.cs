using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReporting
{
    // Simple data entity that represents an item in the source collection.
    public class Item
    {
        public string? Name { get; set; }
        public int Quantity { get; set; }
    }

    // DTO used by the report – contains only the fields required for reporting.
    public class ReportItem
    {
        public string? Name { get; set; }
        public int Quantity { get; set; }
    }

    public class ReportingExample
    {
        /// <summary>
        /// Generates a report by projecting the Name and Quantity fields from the source collection.
        /// </summary>
        /// <param name="items">Source collection containing full item data.</param>
        /// <param name="templatePath">Path to the Aspose.Words template document.</param>
        /// <param name="outputPath">Path where the populated report will be saved.</param>
        public void GenerateReport(IEnumerable<Item> items, string templatePath, string outputPath)
        {
            // Project only the fields required for the report.
            List<ReportItem> reportData = items
                .Select(i => new ReportItem
                {
                    Name = i.Name,
                    Quantity = i.Quantity
                })
                .ToList();

            // Load the template document that contains the reporting placeholders.
            Document doc = new Document(templatePath);

            // Populate the template with the projected data using Aspose.Words ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, reportData);

            // Save the resulting document.
            doc.Save(outputPath);
        }
    }

    // Entry point required for a console application.
    public static class Program
    {
        public static void Main(string[] args)
        {
            // Sample data – in a real scenario this would come from a database or other source.
            var items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 10 },
                new Item { Name = "Banana", Quantity = 5 },
                new Item { Name = "Orange", Quantity = 8 }
            };

            // Paths to the template and output files. Adjust as needed.
            string templatePath = "Template.docx"; // Ensure this file exists.
            string outputPath = "Report.docx";

            var example = new ReportingExample();
            example.GenerateReport(items, templatePath, outputPath);

            Console.WriteLine($"Report generated at '{outputPath}'.");
        }
    }
}
