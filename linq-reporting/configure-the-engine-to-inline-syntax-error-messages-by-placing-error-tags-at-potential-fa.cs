using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model classes
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
        // Note: Price property is intentionally omitted to trigger an inline error.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Simple expression.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // This tag references a non‑existent property and will cause an error.
            builder.Writeln("Missing property: <<[order.NonExistingProperty]>>");

            // Loop over items; the Price property does not exist on Item.
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("Item: <<[item.Name]>> | Qty: <<[item.Quantity]>> | Price: <<[item.Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for reporting.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Quantity = 5 },
                    new Item { Name = "Banana", Quantity = 3 }
                }
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine to inline error messages.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // Build the report. The boolean indicates whether parsing succeeded.
            bool success = engine.BuildReport(reportDoc, order, "order");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);

            // Output the result to the console.
            Console.WriteLine($"Report generation success flag: {success}");
            Console.WriteLine($"Template saved to: {Path.GetFullPath(templatePath)}");
            Console.WriteLine($"Report saved to:   {Path.GetFullPath(reportPath)}");
        }
    }
}
