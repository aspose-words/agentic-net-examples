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
        // Register code page provider (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        string jsonContent = @"
[
    { ""Category"": ""Fruits"", ""Name"": ""Apple"",  ""Price"": 1.20 },
    { ""Category"": ""Fruits"", ""Name"": ""Banana"", ""Price"": 0.80 },
    { ""Category"": ""Vegetables"", ""Name"": ""Carrot"", ""Price"": 0.60 },
    { ""Category"": ""Fruits"", ""Name"": ""Orange"", ""Price"": 1.00 },
    { ""Category"": ""Vegetables"", ""Name"": ""Lettuce"", ""Price"": 1.10 }
]";
        File.WriteAllText(jsonPath, jsonContent.Trim());

        // Deserialize JSON into model.
        List<Item> items = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new();
        ReportModel model = new ReportModel { Items = items };

        // Create the LINQ Reporting template.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin grouping by Category.
        builder.Writeln("<<foreach [catGroup in model.Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[catGroup.Key]>>");
        builder.Writeln("<<foreach [item in catGroup]>>");
        builder.Writeln("- <<[item.Name]>>: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the model.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root model passed to the reporting engine.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual data item.
public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
