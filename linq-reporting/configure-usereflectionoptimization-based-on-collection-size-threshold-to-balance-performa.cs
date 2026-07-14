using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a collection of items.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    public class Program
    {
        // Threshold that decides whether to enable reflection optimization.
        private const int CollectionSizeThreshold = 10;

        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Report of Items:");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Name: <<[item.Name]>>");
            builder.Writeln("Value: <<[item.Value]>>");
            builder.Writeln("<</foreach>>");

            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            // Populate the collection with a variable number of items.
            for (int i = 1; i <= 15; i++)
            {
                model.Items.Add(new Item
                {
                    Name = $"Item {i}",
                    Value = i * 10
                });
            }

            // -----------------------------------------------------------------
            // 4. Configure reflection optimization based on collection size.
            // -----------------------------------------------------------------
            // If the collection is larger than the threshold, enable the optimization;
            // otherwise, disable it to avoid the overhead of dynamic class generation.
            ReportingEngine.UseReflectionOptimization = model.Items.Count > CollectionSizeThreshold;

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
