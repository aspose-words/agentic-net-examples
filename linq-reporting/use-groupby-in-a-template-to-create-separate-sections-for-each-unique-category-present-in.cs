using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
    public int Value { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonContent = @"[
            { ""Name"": ""Apple"",  ""Category"": ""Fruits"",  ""Value"": 5 },
            { ""Name"": ""Banana"", ""Category"": ""Fruits"",  ""Value"": 3 },
            { ""Name"": ""Carrot"", ""Category"": ""Vegetables"", ""Value"": 4 },
            { ""Name"": ""Broccoli"", ""Category"": ""Vegetables"", ""Value"": 6 },
            { ""Name"": ""Chicken"", ""Category"": ""Meat"", ""Value"": 8 }
        ]";

        string jsonPath = "data.json";
        File.WriteAllText(jsonPath, jsonContent);

        // Load JSON into model.
        var items = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new();
        var model = new ReportModel { Items = items };

        // Create template document.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Report grouped by Category");
        builder.Writeln();

        // Grouping using LINQ Reporting tags.
        builder.Writeln("<<foreach [catGroup in Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[catGroup.Key]>>");
        builder.Writeln("<</foreach>>"); // Close outer foreach after inner content.

        // Items within each category.
        builder.Writeln("<<foreach [catGroup in Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("<<foreach [item in catGroup]>>");
        builder.Writeln("- <<[item.Name]>> : <<[item.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save template.
        doc.Save(templatePath);

        // Load template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save final report.
        var outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
