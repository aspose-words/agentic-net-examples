using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

namespace AsposeWordsLinqReportingDemo
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public string Status { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document with LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Begin foreach loop over Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table with a header row.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Status");
            builder.EndRow();

            // Data row – apply background color based on status.
            // Name cell: LightGreen when status is "Completed".
            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Status == \"Completed\"]>>" +
                "<<backColor [\"LightGreen\"]>><<[item.Name]>> <</backColor>><</if>>" +
                "<<if [item.Status != \"Completed\"]>><<[item.Name]>> <</if>>");

            // Status cell: LightYellow when status is "Pending".
            builder.InsertCell();
            builder.Writeln(
                "<<if [item.Status == \"Pending\"]>>" +
                "<<backColor [\"LightYellow\"]>><<[item.Status]>> <</backColor>><</if>>" +
                "<<if [item.Status != \"Pending\"]>><<[item.Status]>> <</if>>");

            builder.EndRow();
            builder.EndTable();

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Prepare sample data.
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Task A", Status = "Completed" },
                    new Item { Name = "Task B", Status = "Pending" },
                    new Item { Name = "Task C", Status = "InProgress" }
                }
            };

            // 3. Load the template and build the report.
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, model, "model");

            // 4. Save the generated report.
            const string outputPath = "Report.docx";
            report.Save(outputPath);
        }
    }
}
