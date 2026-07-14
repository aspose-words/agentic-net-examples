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
    // Simple data model matching the JSON objects.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for temporary files.
        string templatePath = "template.docx";
        string jsonPath = "data.json";

        // Create a large JSON array file.
        const int itemCount = 20000; // Adjust size for benchmarking.
        GenerateLargeJsonFile(jsonPath, itemCount);

        // Create the LINQ Reporting template.
        CreateTemplate(templatePath);

        // Benchmark with reflection optimization enabled.
        ReportingEngine.UseReflectionOptimization = true;
        double timeWithOptimization = RunReport(templatePath, jsonPath, "data");

        // Benchmark with reflection optimization disabled.
        ReportingEngine.UseReflectionOptimization = false;
        double timeWithoutOptimization = RunReport(templatePath, jsonPath, "data");

        // Output the results.
        Console.WriteLine($"Processing time with reflection optimization: {timeWithOptimization:F2} ms");
        Console.WriteLine($"Processing time without reflection optimization: {timeWithoutOptimization:F2} ms");
    }

    // Generates a JSON file containing a large array of Item objects.
    private static void GenerateLargeJsonFile(string filePath, int count)
    {
        var items = new List<Item>(count);
        for (int i = 0; i < count; i++)
        {
            items.Add(new Item { Id = i + 1, Name = $"Item_{i + 1}" });
        }

        string json = JsonConvert.SerializeObject(items);
        File.WriteAllText(filePath, json, Encoding.UTF8);
    }

    // Creates a Word document with a simple foreach tag to iterate over the JSON array.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a header.
        builder.Writeln("Report generated from large JSON array:");
        builder.Writeln();

        // LINQ Reporting foreach tag.
        builder.Writeln("<<foreach [item in data]>>");
        builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Loads the template and JSON data source, builds the report, and returns elapsed milliseconds.
    private static double RunReport(string templatePath, string jsonPath, string rootName)
    {
        // Load fresh template for each run.
        Document doc = new Document(templatePath);

        // Load JSON data source from file.
        using (FileStream jsonStream = File.OpenRead(jsonPath))
        {
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);
            ReportingEngine engine = new ReportingEngine();

            Stopwatch sw = Stopwatch.StartNew();
            engine.BuildReport(doc, jsonDataSource, rootName);
            sw.Stop();

            // Optionally save the generated report (commented out to avoid I/O overhead).
            // doc.Save($"Report_{(ReportingEngine.UseReflectionOptimization ? "Optimized" : "Standard")}.docx");

            return sw.Elapsed.TotalMilliseconds;
        }
    }
}
