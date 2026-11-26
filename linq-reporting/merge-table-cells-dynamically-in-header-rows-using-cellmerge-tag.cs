using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingCellMerge
{
    // Data model classes
    public class ReportModel
    {
        public List<HeaderGroup> Groups { get; set; } = new();
        public List<Item> Items { get; set; } = new();
    }

    public class HeaderGroup
    {
        public string Name { get; set; } = "";
        public int ColumnCount { get; set; }
    }

    public class Item
    {
        public string Col1 { get; set; } = "";
        public string Col2 { get; set; } = "";
        public string Col3 { get; set; } = "";
        public string Col4 { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                Groups = new List<HeaderGroup>
                {
                    new HeaderGroup { Name = "Group A", ColumnCount = 2 },
                    new HeaderGroup { Name = "Group B", ColumnCount = 2 }
                },
                Items = new List<Item>
                {
                    new Item { Col1 = "A1", Col2 = "A2", Col3 = "B1", Col4 = "B2" },
                    new Item { Col1 = "A3", Col2 = "A4", Col3 = "B3", Col4 = "B4" }
                }
            };

            // Create template document
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // ---------- Header table ----------
            Table headerTable = builder.StartTable();

            // First header row – merged group names
            foreach (var group in model.Groups)
            {
                for (int i = 0; i < group.ColumnCount; i++)
                {
                    builder.InsertCell();
                    builder.Write($"<<cellMerge>>{group.Name}");
                }
            }
            builder.EndRow();

            // Second header row – column titles
            int colIndex = 1;
            foreach (var group in model.Groups)
            {
                for (int i = 0; i < group.ColumnCount; i++)
                {
                    builder.InsertCell();
                    builder.Write($"Col {colIndex}");
                    colIndex++;
                }
            }
            builder.EndRow();

            builder.EndTable(); // finish header table

            // ---------- Data table inside foreach ----------
            builder.Writeln("<<foreach [item in Items]>>");

            Table dataTable = builder.StartTable();

            // Data row
            builder.InsertCell();
            builder.Writeln("<<[item.Col1]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Col2]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Col3]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Col4]>>");

            builder.EndRow();
            builder.EndTable();

            builder.Writeln("<</foreach>>");

            // Save the template
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Load the template for reporting
            var doc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the final report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
