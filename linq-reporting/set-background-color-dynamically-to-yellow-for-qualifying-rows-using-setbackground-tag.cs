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

    // Simple item class with an index and a name.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table with a header row.
            var table = builder.StartTable();

            // Header: Index
            builder.InsertCell();
            builder.Writeln("Index");
            // Header: Name
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row: Index column with conditional background.
            builder.InsertCell();
            // If the index is even, apply a yellow background.
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>>" +
                "<<backColor [\"Yellow\"]>><<[item.Index]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>>" +
                "<<[item.Index]>>" +
                "<</if>>");

            // Data row: Name column with the same conditional background.
            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>>" +
                "<<backColor [\"Yellow\"]>><<[item.Name]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>>" +
                "<<[item.Name]>>" +
                "<</if>>");

            builder.EndRow();
            builder.EndTable();

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" },
                    new Item { Index = 4, Name = "Delta" }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(reportPath);
        }
    }
}
