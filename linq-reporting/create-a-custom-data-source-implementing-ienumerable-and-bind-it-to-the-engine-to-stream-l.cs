using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data entity used in the report.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Custom data source that streams a large number of items lazily.
    public class LargeDataSource : IEnumerable<Item>
    {
        private readonly int _count;

        public LargeDataSource(int count = 10000)
        {
            _count = count;
        }

        public IEnumerator<Item> GetEnumerator()
        {
            for (int i = 1; i <= _count; i++)
            {
                // Simulate expensive data retrieval or computation.
                yield return new Item { Id = i, Name = $"Item #{i}" };
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    // Wrapper model that the template will reference.
    public class ReportModel
    {
        public IEnumerable<Item> Items { get; set; } = new List<Item>();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document with LINQ Reporting tags.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Large Data Report ===");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // 2. Load the template for report generation.
            var doc = new Document(templatePath);

            // 3. Prepare the data model with the custom enumerable data source.
            var model = new ReportModel
            {
                Items = new LargeDataSource() // streams 10,000 items lazily.
            };

            // 4. Build the report using the ReportingEngine.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
    }
}
