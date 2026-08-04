using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // File paths.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string jsonPath = Path.Combine(outputDir, "data.json");
        string resultPathOptimized = Path.Combine(outputDir, "result_optimized.docx");
        string resultPathNonOptimized = Path.Combine(outputDir, "result_nonoptimized.docx");

        // Create a large JSON array (e.g., 20,000 items).
        CreateLargeJson(jsonPath, 20000);

        // Create the LINQ Reporting template.
        CreateTemplate(templatePath);

        // Benchmark without reflection optimization.
        ReportingEngine.UseReflectionOptimization = false;
        TimeSpan timeWithout = BuildReport(templatePath, jsonPath, resultPathNonOptimized, "json");

        // Benchmark with reflection optimization enabled.
        ReportingEngine.UseReflectionOptimization = true;
        TimeSpan timeWith = BuildReport(templatePath, jsonPath, resultPathOptimized, "json");

        // Output the measured times.
        Console.WriteLine($"Processing time without reflection optimization: {timeWithout.TotalMilliseconds} ms");
        Console.WriteLine($"Processing time with reflection optimization: {timeWith.TotalMilliseconds} ms");
    }

    // Generates a JSON file containing a large array of items.
    private static void CreateLargeJson(string path, int count)
    {
        var items = new List<Item>();
        for (int i = 1; i <= count; i++)
        {
            items.Add(new Item { Index = i, Name = $"Item {i}" });
        }

        var root = new Root { items = items };
        string json = JsonConvert.SerializeObject(root);
        File.WriteAllText(path, json);
    }

    // Creates a Word template with LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report of items:");
        builder.Writeln("<<foreach [item in json.items]>>");
        builder.Writeln("Index: <<[item.Index]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    // Builds the report, measures processing time, and saves the result.
    private static TimeSpan BuildReport(string templatePath, string jsonPath, string resultPath, string dataSourceName)
    {
        var doc = new Document(templatePath);

        // Ensure the JSON root object is generated so that 'json.items' can be accessed.
        var jsonLoadOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        var jsonDataSource = new JsonDataSource(jsonPath, jsonLoadOptions);

        var engine = new ReportingEngine();

        var stopwatch = Stopwatch.StartNew();
        engine.BuildReport(doc, jsonDataSource, dataSourceName);
        stopwatch.Stop();

        doc.Save(resultPath);
        return stopwatch.Elapsed;
    }

    // Root object for JSON serialization.
    public class Root
    {
        public List<Item> items { get; set; } = new();
    }

    // Item definition used in the JSON array.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }
}
