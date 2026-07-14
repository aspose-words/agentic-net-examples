using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible Unicode handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for temporary files.
        const string jsonPath = "sample.json";
        const string templatePath = "template.docx";
        const string outputPath = "ReportOutput.docx";

        // 1. Create sample JSON with multilingual greetings.
        var sampleData = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alice", Greeting = "Hello (English)" },
                new Item { Name = "Боб", Greeting = "Привет (Russian)" },
                new Item { Name = "陈", Greeting = "你好 (Chinese)" },
                new Item { Name = "ديف", Greeting = "مرحبا (Arabic)" },
                new Item { Name = "ईवा", Greeting = "नमस्ते (Hindi)" }
            }
        };
        // Serialize to JSON file.
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Load JSON into a strongly‑typed model.
        var json = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<ReportModel>(json)!; // Non‑null after deserialization.

        // 3. Build the LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Multilingual Greeting Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Greeting: <<[item.Greeting]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before BuildReport according to rules).
        templateDoc.Save(templatePath);

        // 4. Load the template and generate the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Default options.

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the final report.
        reportDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item containing multilingual text.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Greeting { get; set; } = string.Empty;
}
