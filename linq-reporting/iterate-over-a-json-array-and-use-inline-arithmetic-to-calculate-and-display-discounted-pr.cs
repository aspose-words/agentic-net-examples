using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample JSON data representing a list of products.
        string jsonPath = Path.Combine(outputDir, "products.json");
        string jsonContent = @"
[
    { ""Name"": ""Apple"",  ""Price"": 1.20, ""Discount"": 10 },
    { ""Name"": ""Banana"", ""Price"": 0.80, ""Discount"": 5 },
    { ""Name"": ""Cherry"", ""Price"": 2.50, ""Discount"": 20 }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Product Price Report");
        builder.Writeln("====================================");
        // The data source name will be \"items\" (see BuildReport call below).
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Original Price: $<<[item.Price]>>");
        // Inline arithmetic to calculate discounted price.
        builder.Writeln("Discounted Price: $<<[item.Price * (1 - item.Discount / 100)]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("====================================");

        // 3. Load the JSON data as a JsonDataSource.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple scenario.
        engine.BuildReport(doc, jsonDataSource, "items");

        // 5. Save the generated report.
        string reportPath = Path.Combine(outputDir, "DiscountReport.docx");
        doc.Save(reportPath);
    }
}
