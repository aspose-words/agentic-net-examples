using System;
using System.Collections.Generic;
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
            // Create a blank document and a builder to construct the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table with a header row.
            Table table = builder.StartTable();

            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row with conditional background color based on even index.
            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>><<backColor [\"LightGray\"]>><<[item.Index]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>><<[item.Index]>> <</if>>");

            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Index % 2 == 0]>><<backColor [\"LightGray\"]>><<[item.Name]>> <</backColor>><</if>>" +
                "<<if [item.Index % 2 != 0]>><<[item.Name]>> <</if>>");

            builder.EndRow();

            // Finish the table and the foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            ReportModel model = new()
            {
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Alice" },
                    new Item { Index = 2, Name = "Bob" },
                    new Item { Index = 3, Name = "Charlie" },
                    new Item { Index = 4, Name = "Diana" },
                    new Item { Index = 5, Name = "Eve" }
                }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save("Report.docx");
        }
    }
}
