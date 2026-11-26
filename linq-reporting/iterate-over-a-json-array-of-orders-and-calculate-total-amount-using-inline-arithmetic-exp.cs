using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

// Register code page provider for Aspose.Words.
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class Order
{
    public string CustomerName { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}

public partial class Program
{
    public static void Main()
    {
        // Create sample JSON data file.
        string jsonPath = "orders.json";
        File.WriteAllText(jsonPath,
            @"[
                { ""CustomerName"": ""Alice"", ""Quantity"": 3, ""UnitPrice"": 19.99 },
                { ""CustomerName"": ""Bob"",   ""Quantity"": 5, ""UnitPrice"": 9.50 },
                { ""CustomerName"": ""Carol"", ""Quantity"": 2, ""UnitPrice"": 45.00 }
            ]");

        // Deserialize JSON into a list of Order objects.
        List<Order> orders = JsonConvert.DeserializeObject<List<Order>>(File.ReadAllText(jsonPath)) ?? new();

        // Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        builder.Writeln("Unit Price: <<[order.UnitPrice]>>");
        // Inline arithmetic expression to calculate total per order.
        builder.Writeln("Total: <<[order.Quantity * order.UnitPrice]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, orders, "orders");

        // Save the generated report.
        string reportPath = "report.docx";
        loadedTemplate.Save(reportPath);
    }
}
