using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportGenerator
{
    static void Main()
    {
        // Create temporary files for the template and JSON data.
        string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempDir);

        string templatePath = Path.Combine(tempDir, "Template.docx");
        string jsonPath = Path.Combine(tempDir, "QuarterlyResults.json");
        string outputPath = Path.Combine(tempDir, "ReportWithChart.docx");

        // Build a minimal Word template containing a merge field that references the JSON data.
        // The field name "results" matches the name we will use when building the report.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Quarterly Report");
        // Example merge field that will be replaced by the first item's "Quarter" value.
        builder.InsertField("MERGEFIELD results[0].Quarter");
        builder.Writeln();
        // Example merge field for the "Revenue" value.
        builder.InsertField("MERGEFIELD results[0].Revenue");
        templateDoc.Save(templatePath);

        // Sample JSON data representing quarterly results.
        string jsonContent = @"
        [
            { ""Quarter"": ""Q1"", ""Revenue"": 125000 },
            { ""Quarter"": ""Q2"", ""Revenue"": 150000 },
            { ""Quarter"": ""Q3"", ""Revenue"": 175000 },
            { ""Quarter"": ""Q4"", ""Revenue"": 200000 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create a JSON data source from the temporary file.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report by merging the JSON data into the template.
        // The data source is referenced in the template with the name "results".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "results");

        // Save the populated document.
        doc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}
