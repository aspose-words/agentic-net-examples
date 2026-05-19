using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Public data model for JSON serialization.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    // Wrapper class required by LINQ Reporting to avoid anonymous root objects.
    public class DataWrapper
    {
        public List<Item> Items { get; set; } = new();
    }

    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string templatePath = Path.Combine(outputDir, "template.docx");
        string jsonPath = Path.Combine(outputDir, "data.json");
        string resultWithOpt = Path.Combine(outputDir, "result_with_optimization.docx");
        string resultWithoutOpt = Path.Combine(outputDir, "result_without_optimization.docx");

        // 1. Create a Word template containing LINQ Reporting tags.
        CreateTemplate(templatePath);

        // 2. Generate a large JSON file.
        GenerateLargeJson(jsonPath, itemCount: 20000);

        // 3. Benchmark with reflection optimization enabled.
        long timeWithOpt = BuildReportAndMeasure(
            templatePath,
            jsonPath,
            resultWithOpt,
            useReflectionOptimization: true);

        // 4. Benchmark with reflection optimization disabled.
        long timeWithoutOpt = BuildReportAndMeasure(
            templatePath,
            jsonPath,
            resultWithoutOpt,
            useReflectionOptimization: false);

        // 5. Output the measured times.
        Console.WriteLine($"Processing time with reflection optimization: {timeWithOpt} ms");
        Console.WriteLine($"Processing time without reflection optimization: {timeWithoutOpt} ms");
    }

    // Creates a simple Word document containing a foreach loop over JSON items.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [item in data.Items]>>");
        builder.Writeln("Name: <<[item.Name]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Generates a JSON file with a large array of items.
    private static void GenerateLargeJson(string filePath, int itemCount)
    {
        var wrapper = new DataWrapper();

        for (int i = 0; i < itemCount; i++)
        {
            wrapper.Items.Add(new Item
            {
                Name = $"Item_{i}",
                Value = i
            });
        }

        string json = JsonConvert.SerializeObject(wrapper);
        File.WriteAllText(filePath, json);
    }

    // Builds the report, measures elapsed time, and saves the result document.
    private static long BuildReportAndMeasure(
        string templatePath,
        string jsonPath,
        string resultPath,
        bool useReflectionOptimization)
    {
        // Enable or disable reflection optimization.
        ReportingEngine.UseReflectionOptimization = useReflectionOptimization;

        // Load the template.
        Document doc = new Document(templatePath);

        // Configure JSON loading to keep the wrapper object.
        var jsonOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };

        // Create JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, jsonOptions);

        // Prepare the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Measure the BuildReport execution time.
        Stopwatch sw = Stopwatch.StartNew();
        engine.BuildReport(doc, jsonDataSource, "data");
        sw.Stop();

        // Save the generated document.
        doc.Save(resultPath);

        return sw.ElapsedMilliseconds;
    }
}
