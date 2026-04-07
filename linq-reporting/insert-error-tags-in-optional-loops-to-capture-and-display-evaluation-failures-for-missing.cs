using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingErrorDemo
{
    // Data model for the report.
    public class ReportModel
    {
        // Title of the report.
        public string Title { get; set; } = "Sample Report";

        // Collection that is always present.
        public List<Item> Items { get; set; } = new();

        // Optional collection that may be null to trigger missing‑data handling.
        public List<Item>? OptionalItems { get; set; }
    }

    // Simple item class used in the collections.
    public class Item
    {
        public string Name { get; set; } = "";
        public string? Description { get; set; } // May be null.
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Items =
                {
                    new Item { Name = "Item 1", Description = "First item description" },
                    new Item { Name = "Item 2", Description = null } // Description missing.
                },
                // OptionalItems left null to simulate missing collection.
                OptionalItems = null
            };

            // Create a blank document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write static text and LINQ Reporting tags.
            builder.Writeln("Report Title: <<[model.Title]>>");
            builder.Writeln();

            // Loop over the always‑present collection.
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in model.Items]>>");
            // The <<error>> tag will display any evaluation failure inside the loop.
            builder.Writeln("- <<[item.Name]>>: <<[item.Description]>> <<error>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln();

            // Loop over the optional collection that may be missing.
            builder.Writeln("Optional Items:");
            builder.Writeln("<<foreach [item in model.OptionalItems]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Description]>> <<error>>");
            builder.Writeln("<</foreach>>");

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The bool indicates whether parsing succeeded (relevant when InlineErrorMessages is set).
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated document.
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            // Output the success flag to the console (no interactive prompts).
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to {outputPath}");
        }
    }
}
