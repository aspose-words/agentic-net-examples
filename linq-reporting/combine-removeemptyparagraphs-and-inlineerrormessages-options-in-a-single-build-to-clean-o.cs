using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report";
        // Intentionally missing property to trigger an inline error message.
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the LINQ Reporting template.
            builder.Writeln("<<[model.Title]>>");               // Correct tag.
            builder.Writeln("");                               // Empty paragraph to be removed.
            builder.Writeln("<<[model.MissingProperty]>>");    // Invalid tag – will produce an inline error.
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Save the template (optional, for inspection).
            doc.Save("template.docx");

            // Prepare sample data.
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Quantity = 5 },
                    new Item { Name = "Banana", Quantity = 3 }
                }
            };

            // Configure the reporting engine with both options.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages;

            // Build the report.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("output.docx");

            // Output the result status.
            Console.WriteLine($"Report build success: {success}");
        }
    }
}
