using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            for (int i = 1; i <= 6; i++)
            {
                model.Items.Add(new Item { Index = i, Name = $"Item {i}" });
            }

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Begin a foreach block that iterates over Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Build a simple two‑column table.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row with conditional background colour for even rows.
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

            // Close the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model; the root name is "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            reportDoc.Save(reportPath);
        }
    }
}
