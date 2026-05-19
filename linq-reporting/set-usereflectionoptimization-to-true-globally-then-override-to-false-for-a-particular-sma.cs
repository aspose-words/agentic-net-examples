using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data entity used in the report.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper class that holds a collection of items.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the code page provider is available (required by Aspose.Words on some platforms).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // 1. Set the reflection optimization globally.
            ReportingEngine.UseReflectionOptimization = true;

            // 2. Create a template document with a LINQ Reporting foreach tag.
            const string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // 3. Build a report using a large data source (global optimization stays enabled).
            var largeModel = new ReportModel();
            for (int i = 1; i <= 100; i++)
            {
                largeModel.Items.Add(new Item { Name = $"LargeItem{i}" });
            }

            var largeReport = new Document(templatePath);
            var engineLarge = new ReportingEngine();
            engineLarge.BuildReport(largeReport, largeModel, "model");
            largeReport.Save("LargeReport.docx");

            // 4. Build a report using a small data source with optimization disabled.
            //    Temporarily override the static property for this specific build.
            ReportingEngine.UseReflectionOptimization = false;

            var smallModel = new ReportModel();
            smallModel.Items.Add(new Item { Name = "SmallItemA" });
            smallModel.Items.Add(new Item { Name = "SmallItemB" });

            var smallReport = new Document(templatePath);
            var engineSmall = new ReportingEngine();
            engineSmall.BuildReport(smallReport, smallModel, "model");
            smallReport.Save("SmallReport.docx");

            // 5. Restore the global setting if further processing is required.
            ReportingEngine.UseReflectionOptimization = true;
        }

        // Creates a simple Word document containing a foreach tag that iterates over Items.
        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // The tag iterates over the collection Items in the root object named "model".
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }
    }
}
