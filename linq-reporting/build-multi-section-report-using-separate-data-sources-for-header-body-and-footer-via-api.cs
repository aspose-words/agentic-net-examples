using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingMultiSection
{
    // Data models
    public class ReportModel
    {
        public HeaderModel Header { get; set; } = new();
        public BodyModel Body { get; set; } = new();
        public FooterModel Footer { get; set; } = new();
    }

    public class HeaderModel
    {
        public string Title { get; set; } = "Sample Multi‑Section Report";
        public string Date { get; set; } = DateTime.Now.ToString("yyyy‑MM‑dd");
    }

    public class BodyModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    public class FooterModel
    {
        public string Note { get; set; } = "Report generated with Aspose.Words LINQ Reporting Engine";
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider for any legacy encodings Aspose might need.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel();

            // Body data.
            for (int i = 1; i <= 5; i++)
            {
                model.Body.Items.Add(new Item { Index = i, Name = $"Item {i}" });
            }

            // Create template document.
            var templatePath = "template.docx";
            CreateTemplate(templatePath);

            // Load template for reporting.
            var doc = new Document(templatePath);

            // Build report using LINQ Reporting Engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save final report.
            var outputPath = "report_output.docx";
            doc.Save(outputPath);
        }

        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Header section.
            builder.Writeln("<<[model.Header.Title]>>");
            builder.Writeln("Date: <<[model.Header.Date]>>");
            builder.Writeln(); // Blank line.

            // Body section with a table.
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in model.Body.Items]>>");
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln("<</foreach>>");
            builder.Writeln(); // Blank line.

            // Footer section.
            builder.Writeln("<<[model.Footer.Note]>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
