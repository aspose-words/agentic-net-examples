using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of items to be displayed in the table.
        public List<Item> Items { get; set; } = new();
    }

    // Simple item with an index and a name.
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
            for (int i = 0; i < 10; i++)
            {
                model.Items.Add(new Item { Index = i, Name = $"Item {i}" });
            }

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin the foreach block.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table with a header row.
            Table table = builder.StartTable();

            // Header cells.
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row with conditional background color for even rows.
            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>><<backColor [\"LightGray\"]>><<[item.Index]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>><<[item.Index]>><</if>>");

            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>><<backColor [\"LightGray\"]>><<[item.Name]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>><<[item.Name]>><</if>>");

            builder.EndRow();
            builder.EndTable();

            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
