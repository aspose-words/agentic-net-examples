using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingCellMerge
{
    // Simple data model for the report.
    public class ReportItem
    {
        public string Name { get; set; } = "";
        public int Value { get; set; }
    }

    // Wrapper class that will be passed to the ReportingEngine.
    public class ReportModel
    {
        public List<ReportItem> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document with LINQ Reporting tags.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Open the foreach block before the table.
            builder.Writeln("<<foreach [item in Items]>>");

            // Start the table.
            Table table = builder.StartTable();

            // Header row – two cells that will be merged horizontally using <<cellMerge>>.
            builder.InsertCell();
            builder.Write("<<cellMerge>>Group A");
            builder.InsertCell();
            builder.Write("<<cellMerge>>Group A");
            builder.EndRow();

            // Data row – will be repeated for each item.
            builder.InsertCell();
            builder.Write("<<[item.Name]>>");
            builder.InsertCell();
            builder.Write("<<[item.Value]>>");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Close the foreach block after the table.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<ReportItem>
                {
                    new ReportItem { Name = "Apple",  Value = 10 },
                    new ReportItem { Name = "Banana", Value = 20 },
                    new ReportItem { Name = "Cherry", Value = 30 }
                }
            };

            // 3. Load the template and build the report.
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 4. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
