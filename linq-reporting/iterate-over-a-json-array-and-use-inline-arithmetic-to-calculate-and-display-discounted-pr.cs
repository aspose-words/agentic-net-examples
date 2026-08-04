using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for completeness.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        const string jsonFile = "products.json";
        File.WriteAllText(jsonFile,
            @"[
                { ""Name"": ""Laptop"", ""Price"": 1200.00, ""Discount"": 0.15 },
                { ""Name"": ""Smartphone"", ""Price"": 800.00, ""Discount"": 0.10 },
                { ""Name"": ""Headphones"", ""Price"": 150.00, ""Discount"": 0.20 }
            ]");

        // Create the template document with LINQ Reporting tags.
        const string templateFile = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Price Report");
        builder.Writeln("<<foreach [p in products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Original Price: $<<[p.Price]>>");
        builder.Writeln("Discount: <<[p.Discount * 100]>>%");
        builder.Writeln("Discounted Price: $<<[p.Price * (1 - p.Discount)]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templateFile);

        // Load the template for reporting.
        var doc = new Document(templateFile);

        // Load JSON data source.
        var jsonDataSource = new JsonDataSource(jsonFile);

        // Build the report using the JSON data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "products");

        // Save the final report.
        doc.Save("Report.docx");
    }
}
