using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BenchmarkRemoveEmptyParagraphs
{
    // Data model for the report
    public class Item
    {
        public string Text { get; set; } = string.Empty;
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            const string templatePath = "template.docx";
            const string outputTruePath = "output_remove_true.docx";
            const string outputFalsePath = "output_remove_false.docx";

            // Create the LINQ Reporting template
            CreateTemplate(templatePath);

            // Build a large data source (half of the items are empty)
            var model = new ReportModel();
            const int itemCount = 20000;
            for (int i = 0; i < itemCount; i++)
            {
                model.Items.Add(new Item
                {
                    Text = (i % 2 == 0) ? $"Item {i}" : string.Empty
                });
            }

            // Benchmark with RemoveEmptyParagraphs = true
            TimeSpan timeTrue = RunReport(templatePath, model, true, outputTruePath);

            // Benchmark with RemoveEmptyParagraphs = false
            TimeSpan timeFalse = RunReport(templatePath, model, false, outputFalsePath);

            // Output the results
            Console.WriteLine($"RemoveEmptyParagraphs = true : {timeTrue.TotalMilliseconds} ms");
            Console.WriteLine($"RemoveEmptyParagraphs = false: {timeFalse.TotalMilliseconds} ms");
        }

        // Creates a simple template containing a foreach loop
        private static void CreateTemplate(string path)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<[item.Text]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(path);
        }

        // Runs the report generation and returns the elapsed time
        private static TimeSpan RunReport(string templatePath, ReportModel model, bool removeEmpty, string outputPath)
        {
            var doc = new Document(templatePath);

            var engine = new ReportingEngine();

            // Configure the engine to remove empty paragraphs if requested
            engine.Options = removeEmpty ? ReportBuildOptions.RemoveEmptyParagraphs : ReportBuildOptions.None;

            var stopwatch = Stopwatch.StartNew();
            engine.BuildReport(doc, model, "model");
            stopwatch.Stop();

            doc.Save(outputPath);
            return stopwatch.Elapsed;
        }
    }
}
