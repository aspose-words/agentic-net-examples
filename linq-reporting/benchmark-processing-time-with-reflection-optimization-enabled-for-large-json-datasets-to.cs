using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingBenchmark
{
    public class Item
    {
        public int Id { get; set; } = 0;
        public string Name { get; set; } = "";
        public int Value { get; set; } = 0;
    }

    public class DataRoot
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create template document.
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Report of Items");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>, Value: <<[item.Value]>>");
            builder.Writeln("<</foreach>>");
            doc.Save(templatePath);

            // Load template for reporting.
            var template = new Document(templatePath);

            // Generate large JSON dataset.
            var items = new List<Item>();
            for (int i = 0; i < 5000; i++)
            {
                items.Add(new Item
                {
                    Id = i + 1,
                    Name = $"Item {i + 1}",
                    Value = (i + 1) * 10
                });
            }

            var dataRoot = new DataRoot { Items = items };
            var json = JsonConvert.SerializeObject(dataRoot);
            var dataFromJson = JsonConvert.DeserializeObject<DataRoot>(json)!;

            // Enable reflection optimization.
            ReportingEngine.UseReflectionOptimization = true;

            // Build report and benchmark.
            var engine = new ReportingEngine();
            var stopwatch = Stopwatch.StartNew();
            bool success = engine.BuildReport(template, dataFromJson, "data");
            stopwatch.Stop();

            // Save generated report.
            var reportPath = "report.docx";
            template.Save(reportPath);

            // Output benchmark result.
            Console.WriteLine($"BuildReport succeeded: {success}");
            Console.WriteLine($"Elapsed time (ms): {stopwatch.ElapsedMilliseconds}");
        }
    }
}
