using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encoding scenarios)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a Word template with LINQ Reporting tags
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("-------------------------------------------------");
        // Start a foreach loop over the JSON array named 'orders'
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Date    : <<[order.OrderDate]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount  : <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("-------------------------------------------------");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a JSON file with ISO 8601 date strings
        // -----------------------------------------------------------------
        string jsonPath = Path.Combine(workDir, "Orders.json");
        var orders = new List<Dictionary<string, object>>
        {
            new Dictionary<string, object>
            {
                { "OrderDate", "2023-08-15T14:30:00Z" },
                { "CustomerName", "Alice Johnson" },
                { "Amount", 1250.75 }
            },
            new Dictionary<string, object>
            {
                { "OrderDate", "2023-09-02T09:15:00Z" },
                { "CustomerName", "Bob Smith" },
                { "Amount", 342.00 }
            }
        };
        string jsonContent = System.Text.Json.JsonSerializer.Serialize(orders, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 3. Load the JSON data source with ISO 8601 parsing options
        // -----------------------------------------------------------------
        var jsonLoadOptions = new JsonDataLoadOptions
        {
            // Explicitly specify the ISO 8601 format; the engine also supports it by default.
            ExactDateTimeParseFormats = new List<string> { "yyyy-MM-ddTHH:mm:ssZ" },
            AlwaysGenerateRootObject = true
        };
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, jsonLoadOptions);

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine
        // -----------------------------------------------------------------
        // Reload the template (required before building the report)
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The data source name must match the name used in the template tags ('orders')
        engine.BuildReport(reportDoc, jsonDataSource, "orders");

        // -----------------------------------------------------------------
        // 5. Save the generated report
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
