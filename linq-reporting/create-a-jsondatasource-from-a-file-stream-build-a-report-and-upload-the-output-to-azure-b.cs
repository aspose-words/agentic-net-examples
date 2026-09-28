using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Azure.Storage.Blobs;

public class Program
{
    public static void Main()
    {
        // Register code pages for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths.
        const string jsonPath = "data.json";
        const string templatePath = "template.docx";
        const string reportPath = "Report.docx";

        // Create sample JSON data.
        const string jsonContent = @"{
  ""Title"": ""Sales Report"",
  ""Date"": ""2023-01-01"",
  ""Items"": [
    { ""Name"": ""Product A"", ""Quantity"": 10, ""Price"": 9.99 },
    { ""Name"": ""Product B"", ""Quantity"": 5, ""Price"": 19.99 }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent);

        // Build a Word template with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("Date: <<[model.Date]>>");
        builder.Writeln("");
        builder.Writeln("<<foreach [item in model.Items]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Price]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Load JSON data source from a file stream.
        using (FileStream jsonStream = File.OpenRead(jsonPath))
        {
            var jsonDataSource = new JsonDataSource(jsonStream);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, jsonDataSource, "model");
        }

        // Save the generated report.
        reportDoc.Save(reportPath);

        // Upload the report to Azure Blob Storage (mock implementation).
        try
        {
            const string connectionString = "UseDevelopmentStorage=true";
            const string containerName = "sample-container";
            string blobName = Path.GetFileName(reportPath);

            var blobServiceClient = new BlobServiceClient(connectionString);
            var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
            containerClient.CreateIfNotExists();

            var blobClient = containerClient.GetBlobClient(blobName);
            using (FileStream fileStream = File.OpenRead(reportPath))
            {
                blobClient.Upload(fileStream, overwrite: true);
            }
        }
        catch
        {
            // Swallow any exceptions for this example.
        }
    }
}

// ---------------------------------------------------------------------------
// Mock Azure.Storage.Blobs implementation to allow compilation without the
// real Azure SDK. In a real scenario replace this with the official package.
// ---------------------------------------------------------------------------
namespace Azure.Storage.Blobs
{
    public class BlobServiceClient
    {
        private readonly string _connectionString;
        public BlobServiceClient(string connectionString) => _connectionString = connectionString;
        public BlobContainerClient GetBlobContainerClient(string containerName) => new BlobContainerClient(containerName);
    }

    public class BlobContainerClient
    {
        private readonly string _containerName;
        public BlobContainerClient(string containerName) => _containerName = containerName;
        public void CreateIfNotExists()
        {
            // No‑op for mock; in real code this would create the container.
        }

        public BlobClient GetBlobClient(string blobName) => new BlobClient(blobName);
    }

    public class BlobClient
    {
        private readonly string _blobName;
        public BlobClient(string blobName) => _blobName = blobName;

        public void Upload(Stream content, bool overwrite = false)
        {
            // Simple mock: write the stream to a local file with the blob name.
            using var file = File.Create(_blobName);
            content.CopyTo(file);
        }
    }
}
