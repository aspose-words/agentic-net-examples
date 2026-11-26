using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingElementAtExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data with at least three items
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new() { Name = "First Item" },
                    new() { Name = "Second Item" },
                    new() { Name = "Third Item" },
                    new() { Name = "Fourth Item" }
                }
            };

            // Create a template document programmatically
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            // Insert a LINQ Reporting tag that uses ElementAt to get the third item (index 2)
            builder.Writeln("Third item name: <<[model.Items.ElementAt(2).Name]>>");
            builder.Document.Save(templatePath);

            // Load the template for reporting
            var doc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            doc.Save(outputPath);

            // Indicate completion
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
