using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string json = @"
{
  ""Products"": [
    { ""Name"": ""Apple"",  ""Category"": ""Fruit"",  ""Price"": 0.5 },
    { ""Name"": ""Carrot"", ""Category"": ""Vegetable"", ""Price"": 0.3 },
    { ""Name"": ""Bread"",  ""Category"": ""Bakery"", ""Price"": 1.2 }
  ]
}";
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, json, Encoding.UTF8);

        // Deserialize JSON into model.
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonPath, Encoding.UTF8))!;

        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Header row.
        Table headerTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Category");
        builder.InsertCell(); builder.Writeln("Price");
        builder.EndRow();
        builder.EndTable();

        // Data rows inside a foreach block.
        builder.Writeln("<<foreach [p in Products]>>");
        Table dataTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("<<[p.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Category]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Price]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Build the report using the LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        template.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
    public decimal Price { get; set; }
}
