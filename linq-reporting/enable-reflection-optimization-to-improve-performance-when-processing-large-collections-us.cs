using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the LINQ Reporting template.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            Order order = new Order
            {
                CustomerName = "Acme Corp",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Widget A" },
                    new Item { Index = 2, Name = "Widget B" },
                    new Item { Index = 3, Name = "Widget C" }
                }
            };

            // -----------------------------------------------------------------
            // Step 1: Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a title.
            builder.Writeln("Order Report");
            builder.Writeln();

            // Insert a placeholder for the customer name.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Insert a foreach block to list items.
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Enable reflection optimization for large collections.
            ReportingEngine.UseReflectionOptimization = true;

            ReportingEngine engine = new ReportingEngine();

            // Build the report using the root object name "order".
            engine.BuildReport(reportDoc, order, "order");

            // Save the generated report.
            string reportPath = "Report.docx";
            reportDoc.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }
    }
}
