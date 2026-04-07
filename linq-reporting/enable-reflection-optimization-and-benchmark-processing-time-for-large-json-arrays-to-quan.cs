using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

public class ReflectionOptimizationBenchmark
{
    // Model class matching JSON objects.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        const string templatePath = "template.docx";
        const string jsonPath = "data.json";
        const string outputWithoutOpt = "output_without_optimization.docx";
        const string outputWithOpt = "output_with_optimization.docx";

        // 1. Create the template document programmatically.
        CreateTemplate(templatePath);

        // 2. Generate a large JSON array and save it to a file.
        const int itemCount = 20000; // Adjust size for benchmarking.
        GenerateLargeJson(jsonPath, itemCount);

        // 3. Benchmark without reflection optimization.
        ReportingEngine.UseReflectionOptimization = false;
        long timeWithout = BuildReportAndMeasure(templatePath, jsonPath, outputWithoutOpt);

        // 4. Benchmark with reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;
        long timeWith = BuildReportAndMeasure(templatePath, jsonPath, outputWithOpt);

        // 5. Output the results.
        Console.WriteLine($"Processing time without reflection optimization: {timeWithout} ms");
        Console.WriteLine($"Processing time with reflection optimization:    {timeWith} ms");
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple header.
        builder.Writeln("=== LINQ Reporting Benchmark ===");
        builder.Writeln();

        // LINQ Reporting tags: iterate over JSON items.
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("Name: <<[item.Name]>>\tValue: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(path);
    }

    private static void GenerateLargeJson(string path, int count)
    {
        var items = new List<Item>(count);
        for (int i = 0; i < count; i++)
        {
            items.Add(new Item { Name = $"Item {i}", Value = i });
        }

        string json = JsonConvert.SerializeObject(items);
        File.WriteAllText(path, json);
    }

    private static long BuildReportAndMeasure(string templatePath, string jsonPath, string outputPath)
    {
        // Load the template.
        Document doc = new Document(templatePath);

        // Create JSON data source.
        JsonDataSource jsonSource = new JsonDataSource(jsonPath);

        // Build the report while measuring time.
        ReportingEngine engine = new ReportingEngine();
        Stopwatch sw = Stopwatch.StartNew();
        engine.BuildReport(doc, jsonSource, "items");
        sw.Stop();

        // Save the generated document.
        doc.Save(outputPath);

        return sw.ElapsedMilliseconds;
    }
}
