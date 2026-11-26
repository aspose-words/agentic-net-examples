using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingReflectionBenchmark
{
    public class Item
    {
        public int Id { get; set; } = 0;
        public string Name { get; set; } = string.Empty;
        public double Value { get; set; } = 0.0;
    }

    public class DataRoot
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        static void Main()
        {
            // Ensure code page support for Aspose.Words if needed.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for files.
            const string templatePath = "template.docx";
            const string jsonPath = "data.json";
            const string outputNoOpt = "output_no_optimization.docx";
            const string outputOpt = "output_optimization.docx";

            // Generate large JSON data.
            GenerateJsonData(jsonPath, itemCount: 5000);

            // Create LINQ Reporting template.
            CreateTemplate(templatePath);

            // Load data from JSON.
            var dataRoot = LoadData(jsonPath);

            // Benchmark without reflection optimization.
            ReportingEngine.UseReflectionOptimization = false;
            var timeNoOpt = BuildReport(templatePath, dataRoot, "dataRoot", outputNoOpt);

            // Benchmark with reflection optimization.
            ReportingEngine.UseReflectionOptimization = true;
            var timeOpt = BuildReport(templatePath, dataRoot, "dataRoot", outputOpt);

            // Output benchmark results.
            Console.WriteLine($"BuildReport without reflection optimization: {timeNoOpt.TotalMilliseconds} ms");
            Console.WriteLine($"BuildReport with reflection optimization:    {timeOpt.TotalMilliseconds} ms");
        }

        private static void GenerateJsonData(string filePath, int itemCount)
        {
            var root = new DataRoot();
            for (int i = 1; i <= itemCount; i++)
            {
                root.Items.Add(new Item
                {
                    Id = i,
                    Name = $"Item {i}",
                    Value = i * 1.1
                });
            }

            var json = JsonConvert.SerializeObject(root, Formatting.Indented);
            File.WriteAllText(filePath, json, Encoding.UTF8);
        }

        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Report generated from large JSON data:");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>, Value: <<[item.Value]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }

        private static DataRoot LoadData(string jsonPath)
        {
            var json = File.ReadAllText(jsonPath, Encoding.UTF8);
            return JsonConvert.DeserializeObject<DataRoot>(json) ?? new DataRoot();
        }

        private static TimeSpan BuildReport(string templatePath, DataRoot data, string rootName, string outputPath)
        {
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            var stopwatch = Stopwatch.StartNew();
            engine.BuildReport(doc, data, rootName);
            stopwatch.Stop();

            doc.Save(outputPath);
            return stopwatch.Elapsed;
        }
    }
}
