using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data item used in the report.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    // Custom data source that streams a large number of items lazily.
    public class LargeDataSource : IEnumerable<Item>
    {
        private readonly int _count;

        public LargeDataSource(int count)
        {
            _count = count;
        }

        public IEnumerator<Item> GetEnumerator()
        {
            for (int i = 1; i <= _count; i++)
            {
                // Yield each item on demand – no large in‑memory collection is created.
                yield return new Item
                {
                    Id = i,
                    Name = $"Item {i}",
                    Value = i * 10
                };
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    // Wrapper model that exposes the enumerable to the LINQ Reporting engine.
    public class ReportModel
    {
        public IEnumerable<Item> Items { get; set; } = new List<Item>();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("LINQ Reporting – Streaming Large Data Set");
            builder.Writeln("-------------------------------------------------");
            // foreach tag iterates over the Items collection of the root model.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>, Value: <<[item.Value]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before loading for reporting).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the custom data source and wrapper model.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                // Stream 10,000 items without allocating them all at once.
                Items = new LargeDataSource(10_000)
            };

            // -----------------------------------------------------------------
            // 4. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are needed for this simple example.
            engine.Options = ReportBuildOptions.None;

            // Bind the model to the template. The root name "model" matches the template.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
