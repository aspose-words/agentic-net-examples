using System;
using System.Collections.Generic;
using System.IO;
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
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and generated reports.
            string templatePath = "Template.docx";
            string largeReportPath = "LargeReport.docx";
            string smallReportPath = "SmallReport.docx";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Template uses a foreach loop to iterate over Items and write each Name.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare a large data source.
            // -----------------------------------------------------------------
            ReportModel largeModel = new ReportModel();
            for (int i = 1; i <= 1000; i++)
            {
                largeModel.Items.Add(new Item { Name = $"LargeItem{i}" });
            }

            // -----------------------------------------------------------------
            // 3. Set reflection optimization globally to true.
            // -----------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = true;

            // Load the template for the large report.
            Document largeDoc = new Document(templatePath);
            ReportingEngine largeEngine = new ReportingEngine();

            // Build the report using the large data source.
            largeEngine.BuildReport(largeDoc, largeModel, "model");
            largeDoc.Save(largeReportPath);

            // -----------------------------------------------------------------
            // 4. Prepare a small data source.
            // -----------------------------------------------------------------
            ReportModel smallModel = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "SmallItemA" },
                    new Item { Name = "SmallItemB" }
                }
            };

            // Override reflection optimization to false for this small data source.
            ReportingEngine.UseReflectionOptimization = false;

            // Load the template for the small report.
            Document smallDoc = new Document(templatePath);
            ReportingEngine smallEngine = new ReportingEngine();

            // Build the report using the small data source.
            smallEngine.BuildReport(smallDoc, smallModel, "model");
            smallDoc.Save(smallReportPath);

            // -----------------------------------------------------------------
            // 5. Reset the global setting (optional cleanup).
            // -----------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = true;
        }
    }
}
