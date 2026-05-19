using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data (array of orders)
        string jsonContent = @"
[
    { ""OrderId"": 1001, ""CustomerName"": ""Alice"", ""Quantity"": 3, ""UnitPrice"": 19.99 },
    { ""OrderId"": 1002, ""CustomerName"": ""Bob"",   ""Quantity"": 5, ""UnitPrice"": 9.50 },
    { ""OrderId"": 1003, ""CustomerName"": ""Carol"", ""Quantity"": 2, ""UnitPrice"": 45.00 }
]";
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "orders.json");
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Create a template document with LINQ Reporting tags
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        builder.Writeln("Unit Price: $<<[order.UnitPrice]>>");
        // Inline arithmetic expression to calculate total amount per order
        builder.Writeln("Total: $<<[order.Quantity * order.UnitPrice]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Create a JSON data source from the file
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source name "orders"
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "orders");

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OrdersReport.docx");
        reportDoc.Save(outputPath);
    }
}
