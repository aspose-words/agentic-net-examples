using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data item with a Name property.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    // Root model containing a collection of items.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data with at least five items.
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Item 1" },
                    new Item { Name = "Item 2" },
                    new Item { Name = "Item 3" },
                    new Item { Name = "Item 4" },
                    new Item { Name = "Item 5" }, // Fifth item (index 4)
                    new Item { Name = "Item 6" }
                }
            };

            // Create a template document programmatically.
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that uses ElementAt to fetch the fifth item.
            builder.Writeln("Fifth item: <<[model.Items.ElementAt(4).Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
