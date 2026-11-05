using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required on .NET Core)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
        string jsonContent = @"{
  ""Products"": [
    { ""Name"": ""Apple"",  ""Price"": 0.5,  ""Quantity"": 120 },
    { ""Name"": ""Banana"", ""Price"": 0.3,  ""Quantity"": 200 },
    { ""Name"": ""Orange"", ""Price"": 0.8,  ""Quantity"": 150 }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent);

        // Deserialize JSON into model
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonPath)) ?? new ReportModel();

        // Create template document programmatically
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title
        builder.Writeln("Product Report");
        builder.Writeln();

        // Header row (static)
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();
        builder.EndTable();

        // Data rows using LINQ Reporting foreach
        builder.Writeln("<<foreach [p in Products]>>");
        Table dataTable = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");

        builder.InsertCell();
        builder.Writeln("<<[p.Price]>>");

        builder.InsertCell();
        builder.Writeln("<<[p.Quantity]>>");

        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template
        template.Save(templatePath);

        // Load the template for reporting
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model; root name is "model"
        engine.BuildReport(report, model, "model");

        // Save the final report
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
