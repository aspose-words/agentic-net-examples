using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingConditionalFormatting
{
    // Data model for a single row.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
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
                    new() { Index = 1, Name = "Alpha" },
                    new() { Index = 2, Name = "Beta" },
                    new() { Index = 3, Name = "Gamma" },
                    new() { Index = 4, Name = "Delta" },
                    new() { Index = 5, Name = "Epsilon" }
                }
            };

            // Create the template document.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Begin foreach loop over Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table for each item (header + data row).
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row with conditional background color for even rows.
            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>>" +
                "<<backColor [\"LightGray\"]>><<[item.Index]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>>" +
                "<<[item.Index]>>" +
                "<</if>>");

            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>>" +
                "<<backColor [\"LightGray\"]>><<[item.Name]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>>" +
                "<<[item.Name]>>" +
                "<</if>>");

            builder.EndRow();
            builder.EndTable();

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template for report generation.
            var doc = new Document(templatePath);

            // Build the report.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(doc, model, "model");

            // Save the final report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);

            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
