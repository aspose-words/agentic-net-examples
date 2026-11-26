using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingIsolationExample
{
    // Sample data model.
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    class Program
    {
        static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Step 1: Create the LINQ Reporting template programmatically.
            string templatePath = Path.Combine(outputDir, "template.docx");
            CreateTemplate(templatePath);

            // Simulate two independent user requests.
            ProcessRequest(
                templatePath,
                new Order
                {
                    CustomerName = "Alice",
                    Items = new()
                    {
                        new Item { Index = 1, Name = "Apples" },
                        new Item { Index = 2, Name = "Bananas" }
                    }
                },
                Path.Combine(outputDir, "report_Alice.docx")
            );

            ProcessRequest(
                templatePath,
                new Order
                {
                    CustomerName = "Bob",
                    Items = new()
                    {
                        new Item { Index = 1, Name = "Carrots" },
                        new Item { Index = 2, Name = "Dates" },
                        new Item { Index = 3, Name = "Eggplants" }
                    }
                },
                Path.Combine(outputDir, "report_Bob.docx")
            );
        }

        // Creates a simple Word document containing LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new();
            DocumentBuilder builder = new(doc);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }

        // Processes a single request using an isolated ReportingEngine instance.
        private static void ProcessRequest(string templatePath, Order data, string outputPath)
        {
            // Load the template for this request.
            Document template = new(templatePath);

            // Create a fresh ReportingEngine instance for isolation.
            ReportingEngine engine = new();

            // Build the report using the provided data. The root name must match the tag prefix.
            engine.BuildReport(template, data, "order");

            // Save the generated report.
            template.Save(outputPath);
        }
    }
}
