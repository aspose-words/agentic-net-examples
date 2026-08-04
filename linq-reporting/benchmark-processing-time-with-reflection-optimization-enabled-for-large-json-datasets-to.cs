using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Simple data model for JSON serialization.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    // Wrapper object that holds the collection; required for proper JSON structure.
    public class RootObject
    {
        public List<Item> items { get; set; } = new();
    }

    public static void Main()
    {
        // Paths for temporary files.
        const string jsonPath = "Data.json";
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // 1. Generate a large JSON dataset.
        const int itemCount = 50000; // Adjust for desired size.
        var root = new RootObject();
        for (int i = 0; i < itemCount; i++)
        {
            root.items.Add(new Item { Name = $"Item {i}", Value = i });
        }

        // Serialize to JSON and write to file.
        string json = JsonConvert.SerializeObject(root);
        File.WriteAllText(jsonPath, json);

        // 2. Create a LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated with Aspose.Words LINQ Reporting");
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("<<[item.Name]>> - <<[item.Value]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 3. Load the template.
        var doc = new Document(templatePath);

        // 4. Prepare JSON data source with options.
        var jsonOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        var jsonDataSource = new JsonDataSource(jsonPath, jsonOptions);

        // 5. Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;

        // 6. Build the report and benchmark the processing time.
        var engine = new ReportingEngine();
        var stopwatch = Stopwatch.StartNew();
        engine.BuildReport(doc, jsonDataSource);
        stopwatch.Stop();

        // 7. Save the generated report.
        doc.Save(outputPath);

        // Output the elapsed time.
        Console.WriteLine($"Report generation time with reflection optimization: {stopwatch.ElapsedMilliseconds} ms");
    }
}
