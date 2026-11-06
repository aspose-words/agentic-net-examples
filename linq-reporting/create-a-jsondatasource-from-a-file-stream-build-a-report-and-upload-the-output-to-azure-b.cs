using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

// Register code pages (required by Aspose.Words on some platforms)
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public partial class Program
{
    public static void Main()
    {
        // Prepare sample data model (used only for JSON serialization)
        var model = new ReportModel
        {
            Orders = new()
            {
                new Order
                {
                    CustomerName = "Alice",
                    Items = new()
                    {
                        new Item { Name = "Apple", Quantity = 3 },
                        new Item { Name = "Banana", Quantity = 5 }
                    }
                },
                new Order
                {
                    CustomerName = "Bob",
                    Items = new()
                    {
                        new Item { Name = "Orange", Quantity = 2 },
                        new Item { Name = "Grapes", Quantity = 7 }
                    }
                }
            }
        };

        // Create working directory
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "work");
        Directory.CreateDirectory(workDir);

        // 1. Create the template document programmatically
        string templatePath = Path.Combine(workDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("=== Orders Report ===");
        // Use the root name "model" in the LINQ tags
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 2. Serialize sample data to JSON file
        string jsonPath = Path.Combine(workDir, "data.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(model, Formatting.Indented));

        // 3. Load the template document
        var doc = new Document(templatePath);

        // 4. Create JsonDataSource from file stream and build the report
        using (FileStream jsonStream = new FileStream(jsonPath, FileMode.Open, FileAccess.Read))
        {
            var jsonDataSource = new JsonDataSource(jsonStream);
            var engine = new ReportingEngine();
            // Build the report using root name "model"
            engine.BuildReport(doc, jsonDataSource, "model");
        }

        // 5. Save the generated report
        string reportPath = Path.Combine(workDir, "report.docx");
        doc.Save(reportPath);

        // 6. (Optional) Upload the report to Azure Blob Storage
        // Uncomment and add the Azure.Storage.Blobs NuGet package to enable.
        /*
        string connectionString = Environment.GetEnvironmentVariable("AZURE_STORAGE_CONNECTION_STRING")
                                  ?? "UseDevelopmentStorage=true";
        string containerName = "reports";

        try
        {
            var blobServiceClient = new Azure.Storage.Blobs.BlobServiceClient(connectionString);
            var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
            containerClient.CreateIfNotExists();

            var blobClient = containerClient.GetBlobClient("report.docx");
            using (FileStream uploadStream = File.OpenRead(reportPath))
            {
                blobClient.Upload(uploadStream, overwrite: true);
            }

            Console.WriteLine("Report uploaded successfully to Azure Blob Storage.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Azure Blob upload failed: {ex.Message}");
        }
        */

        Console.WriteLine($"Report generated at: {reportPath}");
    }
}

// Data model classes (public, non‑nullable properties initialized)
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
