using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model for the report.
    public class ReportModel
    {
        // Collection of items to be listed in the report.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item containing a financial amount.
    public class Item
    {
        // Index of the item (for display purposes).
        public int Index { get; set; }

        // Monetary amount that will be rounded in the template.
        public decimal Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Index = 1, Amount = 1234.5678m },
                    new Item { Index = 2, Amount = 9876.5432m },
                    new Item { Index = 3, Amount = 2500.0m }
                }
            };

            // -----------------------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";

            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Title.
            builder.Writeln("Invoice Report");
            builder.Writeln();

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Use System.Math.Round to round the Amount to two decimal places.
            // The static Math type is accessed via the KnownTypes collection of the engine.
            builder.Writeln("Item <<[item.Index]>>: $<<[Math.Round(item.Amount, 2)]>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // Configure the reporting engine.
            var engine = new ReportingEngine();

            // Register System.Math so that its static members can be used in expressions.
            engine.KnownTypes.Add(typeof(Math));

            // Build the report using the model as the data source.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // Step 3: Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
