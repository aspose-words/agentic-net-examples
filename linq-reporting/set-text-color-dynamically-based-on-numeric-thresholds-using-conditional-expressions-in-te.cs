using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Score { get; set; }
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
                    new Item { Name = "Alice", Score = 85 },
                    new Item { Name = "Bob",   Score = 42 },
                    new Item { Name = "Carol", Score = 73 },
                    new Item { Name = "Dave",  Score = 28 }
                }
            };

            // Create the template document programmatically.
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Dynamic Text Color Report");
            builder.Writeln();

            // Begin foreach loop over Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Write item name.
            builder.Writeln("Item: <<[item.Name]>> - ");

            // Apply textColor tag with a conditional expression.
            // Scores below 50 are shown in Red, otherwise in Green.
            builder.Writeln(
                "<<textColor [item.Score < 50 ? \"Red\" : \"Green\"]>>Score: <<[item.Score]>> <</textColor>>");

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            var outputPath = "ReportOutput.docx";
            reportDoc.Save(outputPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine(success
                ? $"Report generated successfully: {Path.GetFullPath(outputPath)}"
                : "Report generation failed.");
        }
    }
}
