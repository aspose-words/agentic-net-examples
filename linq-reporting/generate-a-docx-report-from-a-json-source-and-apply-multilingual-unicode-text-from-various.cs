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
        // Register code page provider for full Unicode support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new ReportModel
        {
            Title = "Multilingual Report",
            Description = "This report contains text in several languages.",
            Items = new List<LocalizedItem>
            {
                new() { Language = "English", Text = "Hello, world!" },
                new() { Language = "Русский", Text = "Привет, мир!" },
                new() { Language = "中文", Text = "你好，世界！" },
                new() { Language = "العربية", Text = "مرحبًا بالعالم!" },
                new() { Language = "हिन्दी", Text = "नमस्ते, दुनिया!" }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Create the LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title and description.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<<[model.Description]>>");
        builder.Writeln();

        // Table header.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Language: <<[item.Language]>>");
        builder.Writeln("Text: <<[item.Text]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Load JSON data into the model.
        string json = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<ReportModel>(json)!;

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Data model matching the JSON structure.
public class ReportModel
{
    public string Title { get; set; } = "";
    public string Description { get; set; } = "";
    public List<LocalizedItem> Items { get; set; } = new();
}

public class LocalizedItem
{
    public string Language { get; set; } = "";
    public string Text { get; set; } = "";
}
