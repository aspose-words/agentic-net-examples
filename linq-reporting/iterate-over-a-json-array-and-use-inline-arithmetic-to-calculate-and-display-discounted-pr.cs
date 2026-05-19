using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some locales).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string dataFile = Path.Combine(workDir, "products.json");
        string templateFile = Path.Combine(workDir, "template.docx");
        string resultFile = Path.Combine(workDir, "DiscountReport.docx");

        // 1. Create sample JSON array.
        string json = @"[
            { ""Name"": ""Apple"",  ""Price"": 10.0, ""Discount"": 0.10 },
            { ""Name"": ""Banana"", ""Price"": 5.0,  ""Discount"": 0.20 },
            { ""Name"": ""Orange"", ""Price"": 8.0,  ""Discount"": 0.15 }
        ]";
        File.WriteAllText(dataFile, json, Encoding.UTF8);

        // 2. Build a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Product Discount Report");
        builder.Writeln("<<foreach [p in products]>>");
        // Display product name, original price, and discounted price calculated inline.
        builder.Writeln("<<[p.Name]>>: Original <<[p.Price]>>  Discounted <<[p.Price * (1 - p.Discount)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (required by the workflow).
        template.Save(templateFile);

        // 3. Load the template (demonstrating the load‑save cycle).
        Document doc = new Document(templateFile);

        // 4. Create a JsonDataSource from the JSON file.
        JsonDataSource jsonData = new JsonDataSource(dataFile);

        // 5. Build the report. The data source name must match the tag used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "products");

        // 6. Save the generated report.
        doc.Save(resultFile);

        Console.WriteLine($"Report generated: {resultFile}");
    }
}
