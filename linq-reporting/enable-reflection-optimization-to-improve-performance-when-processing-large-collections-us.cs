using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportData
    {
        // Collection of items that will be iterated in the template.
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a large collection to demonstrate the benefit of reflection optimization.
            var data = new ReportData();
            for (int i = 1; i <= 10000; i++)
            {
                data.Items.Add(new Item { Id = i, Name = $"Item #{i}" });
            }

            // Build the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a simple foreach tag that iterates over the Items collection.
            builder.Writeln("<<foreach [item in data.Items]>>");
            builder.Writeln("<<[item.Id]>> - <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Enable reflection optimization to speed up member access via runtime proxy generation.
            ReportingEngine.UseReflectionOptimization = true;

            // Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine();
            bool success = engine.BuildReport(template, data, "data");

            // Save the generated report.
            const string outputPath = "Report_ReflectionOptimization.docx";
            template.Save(outputPath);

            // Output a simple confirmation (no interactive input required).
            Console.WriteLine($"Report generated successfully: {success}");
            Console.WriteLine($"Saved to: {outputPath}");
        }
    }
}
