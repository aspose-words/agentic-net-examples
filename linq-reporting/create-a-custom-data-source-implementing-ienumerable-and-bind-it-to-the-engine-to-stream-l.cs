using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data item used in the report.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Custom data source that streams items lazily.
    public class LargeDataSource : IEnumerable<Item>
    {
        private readonly int _count;

        public LargeDataSource(int count) => _count = count;

        public IEnumerator<Item> GetEnumerator()
        {
            for (int i = 0; i < _count; i++)
            {
                // Simulate expensive data retrieval.
                yield return new Item { Index = i, Name = $"Item {i}" };
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    // Wrapper model exposing the collection to the template.
    public class ReportModel
    {
        public IEnumerable<Item> Items { get; set; } = Array.Empty<Item>();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a title.
            builder.Writeln("Large Data Set Report");
            builder.Writeln();

            // Insert a foreach block that iterates over Items.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Index: <<[item.Index]>>  Name: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Prepare the model with a large data source (e.g., 10,000 items).
            ReportModel model = new ReportModel
            {
                Items = new LargeDataSource(10000)
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("LargeDataReport.docx");
        }
    }
}
