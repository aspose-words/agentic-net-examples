using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the root object for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
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

    // External type whose static members will be used inside the template.
    public static class Helper
    {
        // Example static method that formats a string.
        public static string FormatName(string name) => $"**{name.ToUpperInvariant()}**";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -------------------------------------------------
            // 1. Create a template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Header.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Table header.
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[Helper.FormatName(item.Name)]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before BuildReport).
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template back.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Enable reflection optimization and register the external type.
            // -------------------------------------------------
            ReportingEngine.UseReflectionOptimization = true;

            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(Helper));

            // -------------------------------------------------
            // 4. Build the report.
            // -------------------------------------------------
            Order order = new Order(); // Sample data.
            engine.BuildReport(doc, order, "order");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
