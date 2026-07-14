using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // JsonDataSource resides in this namespace
// Note: Azure.Storage.Blobs is omitted because it is not part of the allowed packages.

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings that Aspose.Words might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a simple JSON data file.
        // -----------------------------------------------------------------
        const string jsonFileName = "person.json";
        string jsonContent = @"{
    ""Name"": ""John Doe"",
    ""Age"": 30,
    ""Address"": ""123 Main St, Anytown""
}";
        File.WriteAllText(jsonFileName, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a Word template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags that reference the JSON root object named "person".
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Address: <<[person.Address]>>");

        // Save the template to a temporary file (required before loading it for reporting).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and the JSON data source from a file stream.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        using FileStream jsonStream = File.OpenRead(jsonFileName);
        JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

        // -----------------------------------------------------------------
        // 4. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name used in the template tags is "person".
        engine.BuildReport(doc, jsonDataSource, "person");

        // Save the generated report locally.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);

        // -----------------------------------------------------------------
        // 5. (Optional) Upload the report to Azure Blob Storage.
        // -----------------------------------------------------------------
        // Azure.Storage.Blobs is not included in the allowed package list.
        // In a real scenario you would create a BlobContainerClient and upload the file here.
        // For this self‑contained example we simply inform the user where the report was saved.
        Console.WriteLine($"Report generated and saved to '{reportPath}'.");
        Console.WriteLine("Upload to Azure Blob Storage would be performed here using Azure.Storage.Blobs.");
    }
}
