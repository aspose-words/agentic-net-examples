using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingInlineErrorExample
{
    // Sample data model
    public class ReportModel
    {
        public string Name { get; set; } = "John Doe";

        // This property does NOT exist intentionally to trigger an inline error.
        // public string MissingProperty { get; set; }

        public List<Item> Items { get; set; } = new();

        public ReportModel()
        {
            Items.Add(new Item { Index = 1, Description = "First item" });
            Items.Add(new Item { Index = 2, Description = "Second item" });
        }
    }

    public class Item
    {
        public int Index { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words if needed.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the LINQ Reporting template with tags.
            builder.Writeln("Customer: <<[model.Name]>>");
            // This tag references a non‑existent property and will produce an inline error.
            builder.Writeln("Missing: <<[model.MissingProperty]>>");
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Description]>>");
            builder.Writeln("<</foreach>>");

            // Save the template (optional, just for inspection).
            string templatePath = "template.docx";
            doc.Save(templatePath);

            // Prepare the data model.
            ReportModel model = new ReportModel();

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string outputPath = Path.Combine("output", "report.docx");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
            doc.Save(outputPath);

            // Output the result status.
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
