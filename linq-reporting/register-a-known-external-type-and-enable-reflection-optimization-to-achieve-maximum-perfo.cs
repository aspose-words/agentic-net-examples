using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample data model.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    // Helper class whose static members can be used in the template.
    public static class Helper
    {
        // Returns the upper‑cased version of the supplied text.
        public static string ToUpper(string text) => text?.ToUpperInvariant() ?? string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a foreach block that iterates over model.Items and uses the static helper.
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("Item: <<[Helper.ToUpper(item.Name)]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new() { Name = "Apple" },
                    new() { Name = "Banana" },
                    new() { Name = "Cherry" }
                }
            };

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            // Enable reflection optimization for maximum performance.
            ReportingEngine.UseReflectionOptimization = true;

            var engine = new ReportingEngine();

            // Register the Helper type so its static members can be accessed from the template.
            engine.KnownTypes.Add(typeof(Helper));

            // Load the template document.
            var doc = new Document(templatePath);

            // Build the report. The root object name must match the name used in the template tags ("model").
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
