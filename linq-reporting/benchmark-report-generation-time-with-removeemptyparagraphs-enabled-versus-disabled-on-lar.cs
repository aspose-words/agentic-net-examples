using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBenchmark
{
    // Simple data item used in the report.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        // This field is intentionally left empty for some items to test paragraph removal.
        public string Optional { get; set; } = "";
    }

    // Wrapper class that serves as the root data source for the LINQ Reporting engine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        private const string TemplatePath = "Template.docx";
        private const string OutputNoRemoval = "Report_NoRemoveEmpty.docx";
        private const string OutputWithRemoval = "Report_RemoveEmpty.docx";

        public static void Main()
        {
            // 1. Create the template document programmatically.
            CreateTemplate();

            // 2. Prepare a large data set (e.g., 5,000 items).
            ReportModel model = GenerateData(5000);

            // 3. Benchmark without RemoveEmptyParagraphs option.
            BenchmarkReport(model, ReportBuildOptions.None, OutputNoRemoval, "RemoveEmptyParagraphs disabled");

            // 4. Benchmark with RemoveEmptyParagraphs option enabled.
            BenchmarkReport(model, ReportBuildOptions.RemoveEmptyParagraphs, OutputWithRemoval, "RemoveEmptyParagraphs enabled");
        }

        private static void CreateTemplate()
        {
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Simple title.
            builder.Writeln("LINQ Reporting Benchmark");

            // Begin a foreach block that iterates over Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Paragraph that always contains data.
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");

            // Paragraph that may become empty (Optional can be empty).
            builder.Writeln("<<[item.Optional]>>");

            // End of foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            template.Save(TemplatePath);
        }

        private static ReportModel GenerateData(int count)
        {
            ReportModel model = new ReportModel();

            for (int i = 1; i <= count; i++)
            {
                // Every 10th item has an empty Optional field to trigger empty paragraph removal.
                bool isEmptyOptional = i % 10 == 0;
                model.Items.Add(new Item
                {
                    Index = i,
                    Name = $"Name {i}",
                    Optional = isEmptyOptional ? "" : $"Optional text for item {i}"
                });
            }

            return model;
        }

        private static void BenchmarkReport(ReportModel model, ReportBuildOptions options, string outputPath, string description)
        {
            // Load a fresh copy of the template for each run.
            Document doc = new Document(TemplatePath);

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = options; // Do not use object initializer for the enum.

            // Measure the time taken to build the report.
            Stopwatch sw = Stopwatch.StartNew();
            bool success = engine.BuildReport(doc, model, "model");
            sw.Stop();

            // Save the generated report.
            doc.Save(outputPath);

            // Output the benchmark result.
            Console.WriteLine($"{description}: {sw.ElapsedMilliseconds} ms, success = {success}");
        }
    }
}
