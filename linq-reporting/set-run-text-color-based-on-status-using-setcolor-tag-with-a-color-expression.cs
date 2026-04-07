using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of items to be displayed in the report.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item with a status and a color expression.
    public class Item
    {
        public string Status { get; set; } = "";
        public string Color { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Status = "Completed", Color = "\"Green\"" }, // Color expression can be a string literal.
                    new Item { Status = "Pending",   Color = "\"Orange\"" },
                    new Item { Status = "Failed",    Color = "\"Red\"" }
                }
            };

            // 2. Create the template document programmatically.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report of Items:");
            // Start a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Apply text color based on the item's Color expression and write the status.
            builder.Writeln("<<textColor [item.Color]>><<[item.Status]>> <</textColor>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // 3. Load the template and build the report.
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model; the root name is "model".
            bool success = engine.BuildReport(doc, model, "model");

            // Optionally, you could check the success flag if you enable InlineErrorMessages.
            // For this simple example we just proceed to save the output.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // The program finishes here; no user interaction is required.
        }
    }
}
