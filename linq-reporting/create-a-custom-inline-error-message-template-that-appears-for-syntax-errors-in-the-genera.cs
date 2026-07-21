using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessageExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
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
            // Register code page provider (required for some environments).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string outputPath = "report_with_errors.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Sample Report");

            // Introduce a syntax error by using an unsupported switch "-unknown".
            // The ReportingEngine will insert an inline error message instead of throwing.
            builder.Writeln("<<foreach [item in Items] -unknown>>");
            builder.Writeln("Item name: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var model = new ReportModel();
            model.Items.Add(new Item { Name = "Alpha" });
            model.Items.Add(new Item { Name = "Beta" });

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine to show inline error messages.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The returned flag indicates whether parsing succeeded.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            doc.Save(outputPath);

            // Output the result status.
            Console.WriteLine($"Report generation success flag: {success}");
            Console.WriteLine($"Generated document saved to: {outputPath}");
        }
    }
}
