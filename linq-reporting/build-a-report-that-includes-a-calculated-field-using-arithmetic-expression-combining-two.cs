using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // ---------- 1. Create sample JSON data ----------
        string jsonPath = Path.Combine(workDir, "data.json");
        string jsonContent = @"{ ""Price"": 12.5, ""Quantity"": 4 }";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // ---------- 2. Create the template document ----------
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("Price   : <<[model.Price]>>");
        builder.Writeln("Quantity: <<[model.Quantity]>>");
        // Calculated field: Price * Quantity
        builder.Writeln("Total   : <<[model.Price * model.Quantity]>>");

        // Save the template so it can be loaded later (required by the lifecycle rule).
        templateDoc.Save(templatePath);

        // ---------- 3. Load the template ----------
        Document doc = new Document(templatePath);

        // ---------- 4. Create JSON data source ----------
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // ---------- 5. Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        // The root object name used in the template tags is "model".
        engine.BuildReport(doc, jsonDataSource, "model");

        // ---------- 6. Save the generated report ----------
        string reportPath = Path.Combine(workDir, "report.docx");
        doc.Save(reportPath);

        Console.WriteLine("Report generated at: " + reportPath);
    }
}
