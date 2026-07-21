using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model for the report.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();
    }

    // Simple item class used in the collection.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a template document programmatically.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report for a large data set:");
            // LINQ Reporting foreach tag.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // 2. Load the template document.
            Document doc = new Document(templatePath);

            // 3. Prepare a large data set.
            ReportModel model = new ReportModel();
            const int itemCount = 10000; // Simulate a large collection.
            for (int i = 1; i <= itemCount; i++)
            {
                model.Items.Add(new Item
                {
                    Index = i,
                    Name = $"Item_{i}"
                });
            }

            // 4. Enable reflection optimization (engine caching) for faster processing.
            ReportingEngine.UseReflectionOptimization = true;

            // 5. Build the report.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this scenario.
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine($"Report generated successfully at: {reportPath}");
        }
    }
}
