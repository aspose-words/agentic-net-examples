using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingCustomDataSource
{
    // Simple data item.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public double Value { get; set; }
    }

    // Custom enumerable data source that streams items lazily.
    public class LargeDataSource : IEnumerable<Item>
    {
        private readonly int _count;

        public LargeDataSource(int count) => _count = count;

        public IEnumerator<Item> GetEnumerator()
        {
            for (int i = 1; i <= _count; i++)
            {
                yield return new Item
                {
                    Id = i,
                    Name = $"Item {i}",
                    Value = i * 1.1
                };
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    // Wrapper model required by the reporting engine.
    public class ReportModel
    {
        public IEnumerable<Item> Items { get; set; } = new List<Item>();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Report generated from a custom IEnumerable data source");
            builder.Writeln();

            // Begin foreach block.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table that will be repeated for each item.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Id");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Value");
            builder.EndRow();

            // Data row – placeholders will be replaced by the engine.
            builder.InsertCell();
            builder.Writeln("<<[item.Id]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Value]>>");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // End foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and generate the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // Prepare the large data source and wrapper model.
            var dataSource = new LargeDataSource(1000);
            var model = new ReportModel { Items = dataSource };

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
