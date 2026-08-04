using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample JSON data representing an array of orders.
        // -----------------------------------------------------------------
        string jsonPath = "orders.json";
        string jsonContent = @"
[
  { ""OrderId"": 1, ""CustomerName"": ""Alice"", ""Quantity"": 3, ""UnitPrice"": 19.99 },
  { ""OrderId"": 2, ""CustomerName"": ""Bob"",   ""Quantity"": 5, ""UnitPrice"": 9.50 },
  { ""OrderId"": 3, ""CustomerName"": ""Carol"", ""Quantity"": 2, ""UnitPrice"": 45.00 }
]
";
        File.WriteAllText(jsonPath, jsonContent.Trim());

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("------------------------------");

        // Begin a foreach loop over the JSON array named 'orders'.
        builder.Writeln("<<foreach [order in orders]>>");

        // Output order details and calculate line total using an inline expression.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        builder.Writeln("Unit Price: $<<[order.UnitPrice]>>");
        builder.Writeln("Line Total: $<<[order.Quantity * order.UnitPrice]>>");
        builder.Writeln(""); // Blank line for readability.

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the JSON data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Default options.
        engine.BuildReport(reportDoc, jsonData, "orders");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
