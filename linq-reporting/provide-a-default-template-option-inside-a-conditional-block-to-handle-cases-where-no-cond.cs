using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalDefault
{
    // Data model for the report.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item with flags used in conditional logic.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public bool IsA { get; set; }
        public bool IsB { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Condition 1: item belongs to group A.
            builder.Writeln("<<if [item.IsA]>>Item A: <<[item.Name]>> <</if>>");

            // Condition 2: item belongs to group B.
            builder.Writeln("<<if [item.IsB]>>Item B: <<[item.Name]>> <</if>>");

            // Default case when neither condition is true.
            builder.Writeln("<<if [!item.IsA && !item.IsB]>>Item Other: <<[item.Name]>> <</if>>");

            // End of the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and prepare sample data.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Alpha", IsA = true },
                    new Item { Name = "Beta", IsB = true },
                    new Item { Name = "Gamma" } // Neither IsA nor IsB.
                }
            };

            // -------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);

            // Optional: indicate completion (no interactive input).
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
