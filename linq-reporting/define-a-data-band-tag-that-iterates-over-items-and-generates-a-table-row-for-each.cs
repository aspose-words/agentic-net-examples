using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingExample
{
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Sample data.
            var model = new ReportModel
            {
                Items = new()
                {
                    new Item { Index = 1, Name = "Apple" },
                    new Item { Index = 2, Name = "Banana" },
                    new Item { Index = 3, Name = "Cherry" }
                }
            };

            // Output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // Build the template document.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "template.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Items Report");
            builder.Writeln(); // empty line

            // Header table (static, not repeated).
            Table headerTable = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();
            builder.EndTable();

            // Data rows – each iteration creates a one‑row table.
            builder.Writeln("<<foreach [item in Items]>>");
            Table dataTable = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // Build the report.
            // -----------------------------------------------------------------
            var templateDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // Save the generated report.
            string reportPath = Path.Combine(outputDir, "report.docx");
            templateDoc.Save(reportPath);

            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
