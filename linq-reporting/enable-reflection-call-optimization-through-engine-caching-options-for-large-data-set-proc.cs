using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of items to be iterated in the template.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item displayed in the report.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare a large data set (e.g., 10,000 items).
            // -----------------------------------------------------------------
            var model = new ReportModel();

            const int itemCount = 10000;
            for (int i = 1; i <= itemCount; i++)
            {
                model.Items.Add(new Item
                {
                    Index = i,
                    Name = $"Item #{i}"
                });
            }

            // -----------------------------------------------------------------
            // 2. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Add a simple heading.
            builder.Writeln("Large Data Set Report");
            builder.Writeln();

            // Insert a foreach tag that will iterate over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Each line will show the index and name of the item.
            builder.Writeln("<<[item.Index]>> - <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Enable reflection optimization (engine caching) for the reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = true;

            // -----------------------------------------------------------------
            // 4. Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
