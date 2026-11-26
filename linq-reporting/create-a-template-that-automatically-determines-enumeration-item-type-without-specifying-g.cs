using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDemo
{
    // Sample data model.
    public class Order
    {
        public int OrderId { get; set; } = 1001;
        public List<Item> Items { get; set; } = new()
        {
            new Item { Index = 1, Name = "Apple" },
            new Item { Index = 2, Name = "Banana" },
            new Item { Index = 3, Name = "Cherry" }
        };
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple title.
            builder.Writeln("Order Report");
            builder.Writeln("Order ID: <<[order.OrderId]>>");
            builder.Writeln();

            // Foreach loop without specifying generic type parameters.
            // The engine will infer the item type (Item) at runtime.
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Prepare the root data object.
            Order order = new();

            // Build the report. The root name must match the tag prefix used in the template.
            engine.BuildReport(reportDoc, order, "order");

            // Save the generated report.
            reportDoc.Save(outputPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
