using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineIsolationExample
{
    // Simple data model for the report.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    public class Program
    {
        // Paths for the template and generated reports.
        private const string TemplatePath = "Template.docx";
        private const string Report1Path = "Report1.docx";
        private const string Report2Path = "Report2.docx";

        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some locales).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // Step 1: Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            CreateTemplate();

            // -----------------------------------------------------------------
            // Step 2: Simulate two independent user requests.
            // Each request gets its own ReportingEngine instance.
            // -----------------------------------------------------------------
            // First request data.
            Order order1 = new Order
            {
                CustomerName = "Alice Johnson",
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Quantity = 5 },
                    new Item { Name = "Banana", Quantity = 12 }
                }
            };
            GenerateReport(order1, Report1Path);

            // Second request data.
            Order order2 = new Order
            {
                CustomerName = "Bob Smith",
                Items = new List<Item>
                {
                    new Item { Name = "Orange", Quantity = 8 },
                    new Item { Name = "Grapes", Quantity = 3 },
                    new Item { Name = "Mango", Quantity = 4 }
                }
            };
            GenerateReport(order2, Report2Path);
        }

        // Creates a Word document containing LINQ Reporting tags and saves it as the template.
        private static void CreateTemplate()
        {
            // Create a blank document.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write static text and LINQ Reporting tags.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Ordered Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Name]>> (Qty: <<[item.Quantity]>>)");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(TemplatePath);
        }

        // Generates a report for a given order using an isolated ReportingEngine instance.
        private static void GenerateReport(Order order, string outputPath)
        {
            // Load the template document.
            Document doc = new Document(TemplatePath);

            // Create a new ReportingEngine for this request.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The root object name must match the tag prefix used in the template.
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
