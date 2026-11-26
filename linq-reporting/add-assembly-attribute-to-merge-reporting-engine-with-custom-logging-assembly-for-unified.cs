using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        public string Name { get; set; } = "John Doe";

        public List<Item> Items { get; set; } = new()
        {
            new Item { Index = 1, Description = "Item One" },
            new Item { Index = 2, Description = "Item Two" }
        };
    }

    public class Item
    {
        public int Index { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public static class Program
    {
        public static void Main()
        {
            // Prepare working directories.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a template document programmatically.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Insert a simple placeholder that will be replaced by the model's Name property.
            builder.Writeln("Customer Name: <<[model.Name]>>");
            builder.Writeln();

            // Insert a table that iterates over Items collection.
            builder.Writeln("<<foreach [item in model.Items]>>");
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Description");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Description]>>");
            builder.EndRow();

            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            Document doc = new(templatePath);

            // Create the reporting engine and configure options.
            ReportingEngine engine = new();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report using the model as the root object.
            ReportModel model = new();
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string resultPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(resultPath);

            // Output simple status.
            Console.WriteLine(success
                ? $"Report generated successfully: {resultPath}"
                : "Report generation completed with errors. See inline messages in the document.");
        }
    }
}
