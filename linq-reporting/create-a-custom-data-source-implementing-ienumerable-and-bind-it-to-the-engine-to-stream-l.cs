using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data entity used in the report.
    public class Record
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    // Custom data source that streams a large number of records lazily.
    public class LargeDataSource : IEnumerable<Record>
    {
        private readonly int _count;

        public LargeDataSource(int count)
        {
            _count = count;
        }

        public IEnumerator<Record> GetEnumerator()
        {
            // Yield records one by one to avoid loading everything into memory at once.
            for (int i = 1; i <= _count; i++)
            {
                yield return new Record
                {
                    Id = i,
                    Name = $"Item {i}",
                    Value = i * 10
                };
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    // Wrapper class that exposes the collection to the LINQ Reporting engine.
    public class ReportModel
    {
        public IEnumerable<Record> Records { get; set; } = new List<Record>();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document programmatically.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Large Data Report ===");
            // LINQ Reporting foreach tag iterating over the Records collection.
            builder.Writeln("<<foreach [rec in Records]>>");
            builder.Writeln("Id: <<[rec.Id]>>, Name: <<[rec.Name]>>, Value: <<[rec.Value]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // 2. Load the template for report generation.
            var doc = new Document(templatePath);

            // 3. Prepare the data model with a streaming data source.
            var model = new ReportModel
            {
                // Stream 10,000 records without holding them all in memory.
                Records = new LargeDataSource(10_000)
            };

            // 4. Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
