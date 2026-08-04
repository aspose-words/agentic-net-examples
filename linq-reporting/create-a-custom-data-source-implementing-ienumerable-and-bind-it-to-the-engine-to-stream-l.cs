using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data entity used in the report.
    public class ReportItem
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper model that exposes the custom data source to the reporting engine.
    public class ReportModel
    {
        public IEnumerable<ReportItem> Items { get; set; } = new List<ReportItem>();
    }

    // Custom data source that streams items on demand.
    public class LargeDataSource : IEnumerable<ReportItem>
    {
        private readonly int _count;

        public LargeDataSource(int count) => _count = count;

        public IEnumerator<ReportItem> GetEnumerator()
        {
            // Simulate streaming a large data set without materialising it all at once.
            for (int i = 1; i <= _count; i++)
            {
                yield return new ReportItem
                {
                    Index = i,
                    Name = $"Item #{i}"
                };
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document that will serve as the template.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Add a title.
            builder.Writeln("Large Data Set Report");
            builder.Writeln();

            // Begin the foreach block that iterates over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a simple table to display each item's data.
            var table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row – the engine will repeat this row for each item.
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();

            // End the table and the foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Prepare the model with the custom streaming data source.
            var model = new ReportModel
            {
                Items = new LargeDataSource(1000) // Stream 1,000 items.
            };

            // Build the report.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // Default options.
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("LargeDataReport.docx");
        }
    }
}
