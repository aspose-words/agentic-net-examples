using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingArithmeticExample
{
    // Data model classes
    public class Order
    {
        // Initialize to avoid nullable warnings
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public decimal Price { get; set; }
        public int Quantity { get; set; }

        // Calculated property used in the template (LINQ Reporting does not support inline arithmetic)
        public decimal Total => Price * Quantity;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Write LINQ Reporting tags into the template
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Price: <<[item.Price]>>");
            builder.Writeln("Quantity: <<[item.Quantity]>>");
            // Use the calculated property instead of an inline expression
            builder.Writeln("Total: <<[item.Total]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Prepare sample data
            // -------------------------------------------------
            var order = new Order
            {
                Items = new List<Item>
                {
                    new Item { Price = 19.99m, Quantity = 2 },
                    new Item { Price = 5.50m,  Quantity = 5 },
                    new Item { Price = 12.30m, Quantity = 1 }
                }
            };

            // -------------------------------------------------
            // 3. Load the template and build the report
            // -------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the root object name "order"
            engine.BuildReport(doc, order, "order");

            // -------------------------------------------------
            // 4. Save the generated report
            // -------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
