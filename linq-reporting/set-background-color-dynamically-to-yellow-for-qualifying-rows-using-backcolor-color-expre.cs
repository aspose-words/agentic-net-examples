using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingBackColorExample
{
    public class Item
    {
        public string Name { get; set; } = "";
        public int Value { get; set; }
        public bool IsHighlighted { get; set; }
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words if needed.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create template document.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Items Report");
            builder.Writeln("<<foreach [item in Items]>>");

            // Start table.
            var table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Value");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            builder.Writeln("<<backColor [item.IsHighlighted ? \"Yellow\" : \"White\"]>><<[item.Name]>> <</backColor>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Value]>>");
            builder.EndRow();

            // End table.
            builder.EndTable();

            builder.Writeln("<</foreach>>");

            // Save template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load template for reporting.
            var doc = new Document(templatePath);

            // Sample data.
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Item 1", Value = 100, IsHighlighted = true },
                    new Item { Name = "Item 2", Value = 200, IsHighlighted = false },
                    new Item { Name = "Item 3", Value = 300, IsHighlighted = true }
                }
            };

            // Build report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save final report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
