using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data: an array of orders, each with a list of items.
        string json = @"
[
  {
    ""Id"": 1,
    ""CustomerName"": ""Alice"",
    ""Items"": [
      { ""ProductName"": ""Widget"", ""Quantity"": 2, ""UnitPrice"": 10.5 },
      { ""ProductName"": ""Gadget"", ""Quantity"": 1, ""UnitPrice"": 20.0 }
    ]
  },
  {
    ""Id"": 2,
    ""CustomerName"": ""Bob"",
    ""Items"": [
      { ""ProductName"": ""Thing"", ""Quantity"": 5, ""UnitPrice"": 3.0 }
    ]
  }
]";

        // Create a template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a foreach loop over the orders array.
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        // Inline arithmetic expression calculates total amount for each order.
        builder.Writeln("Total Amount: <<[order.Items.Sum(i => i.Quantity * i.UnitPrice)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template back for reporting.
        Document reportDoc = new Document(templatePath);

        // Create a JSON data source from the JSON string using a memory stream.
        using MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // The root name "orders" matches the array name used in the template tags.
        engine.BuildReport(reportDoc, jsonDataSource, "orders");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
