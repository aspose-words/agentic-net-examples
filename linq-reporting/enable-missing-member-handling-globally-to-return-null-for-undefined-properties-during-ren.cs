using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingMissingMemberDemo
{
    // Simple data model with a collection.
    public class Order
    {
        public int Id { get; set; }
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create a template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Write some text and LINQ Reporting tags.
            builder.Writeln("Order ID: <<[order.Id]>>");
            // This property does NOT exist in the Order class.
            builder.Writeln("Missing property (should be null): <<[order.MissingProp]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("  Name: <<[item.Name]>>");
            // This property does NOT exist in the Item class.
            builder.Writeln("  Missing: <<[item.NonExisting]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template document.
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare sample data.
            // -------------------------------------------------
            var order = new Order
            {
                Id = 123,
                Items = new List<Item>
                {
                    new Item { Name = "Apple" },
                    new Item { Name = "Banana" }
                }
            };

            // -------------------------------------------------
            // 4. Configure the ReportingEngine to treat missing members as null.
            // -------------------------------------------------
            var engine = new ReportingEngine
            {
                // AllowMissingMembers makes undefined members evaluate to null.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Build the report using the root object name "order".
            engine.BuildReport(doc, order, "order");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
