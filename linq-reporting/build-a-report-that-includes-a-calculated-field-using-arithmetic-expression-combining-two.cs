using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "order.json");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create the JSON data source file with two numeric properties.
        // -----------------------------------------------------------------
        string jsonContent = @"{ ""Price"": 9.99, ""Quantity"": 5 }";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build the Word template programmatically and insert LINQ tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Price   : <<[order.Price]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        // Calculated field: multiplication of Price and Quantity
        builder.Writeln("Total   : <<[order.Price * order.Quantity]>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the JSON data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple example
        engine.BuildReport(reportDoc, jsonDataSource, "order");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);

        // Optional: indicate completion (no interactive prompts)
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
