using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Item
    {
        public string Name { get; set; } = "";
    }

    // Wrapper for a collection of items – used for the large data source.
    public class LargeData
    {
        public List<Item> Items { get; set; } = new();
    }

    // Wrapper for a collection of items – used for the small data source.
    public class SmallData
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Set the reflection optimization globally.
            ReportingEngine.UseReflectionOptimization = true;

            // -----------------------------------------------------------------
            // 1. Create a template document with a simple foreach tag.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a heading.
            builder.Writeln("Report of Items:");
            // Insert the LINQ Reporting foreach tag.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Build a report using a large data source (global optimization stays true).
            // -----------------------------------------------------------------
            LargeData largeData = new LargeData();
            for (int i = 1; i <= 100; i++)
            {
                largeData.Items.Add(new Item { Name = $"Item {i}" });
            }

            // Load the template and build the report.
            Document largeReport = new Document(templatePath);
            ReportingEngine engineLarge = new ReportingEngine();
            engineLarge.BuildReport(largeReport, largeData, "data");
            largeReport.Save("LargeReport.docx");

            // -----------------------------------------------------------------
            // 3. Override the reflection optimization for a small data source.
            // -----------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = false; // Override for this specific report.

            SmallData smallData = new SmallData();
            smallData.Items.Add(new Item { Name = "Alpha" });
            smallData.Items.Add(new Item { Name = "Beta" });
            smallData.Items.Add(new Item { Name = "Gamma" });

            // Load the same template again and build the report.
            Document smallReport = new Document(templatePath);
            ReportingEngine engineSmall = new ReportingEngine();
            engineSmall.BuildReport(smallReport, smallData, "data");
            smallReport.Save("SmallReport.docx");

            // (Optional) Restore the global setting if further processing is required.
            ReportingEngine.UseReflectionOptimization = true;
        }
    }
}
