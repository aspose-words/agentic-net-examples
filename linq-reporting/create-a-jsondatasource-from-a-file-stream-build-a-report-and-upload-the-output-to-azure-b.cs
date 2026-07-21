using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonFilePath = Path.Combine(Directory.GetCurrentDirectory(), "persons.json");
        string jsonContent = @"[
            { ""Name"": ""John Doe"", ""Age"": 30 },
            { ""Name"": ""Jane Smith"", ""Age"": 25 }
        ]";
        File.WriteAllText(jsonFilePath, jsonContent);

        // Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template back (required before building the report).
        Document reportDoc = new Document(templatePath);

        // Load JSON data from a file stream.
        using (FileStream jsonStream = File.OpenRead(jsonFilePath))
        {
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, jsonDataSource, "persons");
        }

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");
        reportDoc.Save(reportPath);

        // NOTE: Azure Blob Storage upload requires the Azure.Storage.Blobs package,
        // which is not part of the required package list. To keep the example
        // self‑contained and compilable, the upload step is represented as a placeholder.
        Console.WriteLine($"Report generated at: {reportPath}");
        Console.WriteLine("Upload to Azure Blob Storage would be performed here.");
    }
}
