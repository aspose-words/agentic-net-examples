using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingReflectionOptimization
{
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public double Value { get; set; }
    }

    public class Root
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        private const string TemplateFileName = "template.docx";
        private const string JsonFileName = "data.json";
        private const string OutputNonOptimized = "output_nonoptimized.docx";
        private const string OutputOptimized = "output_optimized.docx";

        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare large JSON data.
            Root data = GenerateSampleData(10000);
            string json = JsonConvert.SerializeObject(data);
            File.WriteAllText(JsonFileName, json);

            // Load data back from JSON.
            Root root = JsonConvert.DeserializeObject<Root>(File.ReadAllText(JsonFileName))!;

            // Create and save the template document.
            CreateTemplate();

            // Benchmark without reflection optimization.
            TimeSpan nonOptimizedTime = RunReport(root, false, OutputNonOptimized);
            Console.WriteLine($"Report without reflection optimization: {nonOptimizedTime.TotalMilliseconds} ms");

            // Benchmark with reflection optimization.
            TimeSpan optimizedTime = RunReport(root, true, OutputOptimized);
            Console.WriteLine($"Report with reflection optimization: {optimizedTime.TotalMilliseconds} ms");
        }

        private static Root GenerateSampleData(int count)
        {
            var root = new Root();
            for (int i = 1; i <= count; i++)
            {
                root.Items.Add(new Item
                {
                    Id = i,
                    Name = $"Item {i}",
                    Value = i * 0.1
                });
            }
            return root;
        }

        private static void CreateTemplate()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>, Value: <<[item.Value]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(TemplateFileName);
        }

        private static TimeSpan RunReport(Root data, bool useOptimization, string outputPath)
        {
            ReportingEngine.UseReflectionOptimization = useOptimization;

            Document template = new Document(TemplateFileName);
            ReportingEngine engine = new ReportingEngine();

            Stopwatch sw = Stopwatch.StartNew();
            engine.BuildReport(template, data, "root");
            sw.Stop();

            template.Save(outputPath);
            return sw.Elapsed;
        }
    }
}
