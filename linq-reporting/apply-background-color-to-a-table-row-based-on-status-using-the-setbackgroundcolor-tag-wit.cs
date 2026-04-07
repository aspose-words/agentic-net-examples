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

    // Individual item with a name and a status.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create sample data.
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Task A", Status = "Completed" },
                    new Item { Name = "Task B", Status = "InProgress" },
                    new Item { Name = "Task C", Status = "Completed" },
                    new Item { Name = "Task D", Status = "Pending" }
                }
            };

            // 2. Build the template document programmatically.
            var templatePath = "Template.docx";
            var outputPath = "Report.docx";

            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Begin the foreach loop over Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table for each iteration (header + data row).
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Status");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            // Insert the name field.
            builder.Writeln("<<[item.Name]>>");

            builder.InsertCell();
            // Conditional background color: LightGreen for Completed status.
            builder.Writeln(
                "<<if [item.Status == \"Completed\"]>><<backColor [\"LightGreen\"]>><<[item.Status]>> <</backColor>><</if>>" +
                "<<if [item.Status != \"Completed\"]>><<[item.Status]>> <</if>>");

            builder.EndRow();

            // Finish the table and the foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // 3. Load the template and build the report.
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 4. Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
