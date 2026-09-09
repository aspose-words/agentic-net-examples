using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create sample JSON data representing an array of orders.
        // Each order has an Id, CustomerName, Quantity and UnitPrice.
        string jsonContent = @"
[
  { ""OrderId"": 1, ""CustomerName"": ""Alice"", ""Quantity"": 3, ""UnitPrice"": 19.99 },
  { ""OrderId"": 2, ""CustomerName"": ""Bob"",   ""Quantity"": 5, ""UnitPrice"": 9.50 },
  { ""OrderId"": 3, ""CustomerName"": ""Carol"", ""Quantity"": 2, ""UnitPrice"": 45.00 }
]";
        string jsonPath = Path.Combine(workDir, "orders.json");
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document programmatically.
        // The template will iterate over the JSON array and calculate total amount per order.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header
        builder.Writeln("Order Report");
        builder.Writeln("------------------------------");

        // Begin foreach over the JSON data source named \"orders\".
        builder.Writeln("<<foreach [order in orders]>>");
        // Output order details and calculate total = Quantity * UnitPrice using an inline expression.
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        builder.Writeln("Unit Price: $<<[order.UnitPrice]>>");
        builder.Writeln("Total Amount: $<<[order.Quantity * order.UnitPrice]>>");
        builder.Writeln(""); // Blank line between orders
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document for report generation.
        Document reportDoc = new Document(templatePath);

        // 4. Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name used in the template is \"orders\".
        engine.BuildReport(reportDoc, jsonDataSource, "orders");

        // 6. Save the generated report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion (no interactive prompts as required).
        Console.WriteLine("Report generated successfully at: " + reportPath);
    }
}
