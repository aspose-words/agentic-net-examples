using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Order
{
    public int OrderId { get; set; } = 0;
    public string CustomerName { get; set; } = "";
    public List<OrderItem> Items { get; set; } = new();
}

public class OrderItem
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; } = 0;
    public decimal UnitPrice { get; set; } = 0m;
}

public class RootData
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data.
        string json = @"
[
  {
    ""OrderId"": 1001,
    ""CustomerName"": ""Alice Smith"",
    ""Items"": [
      { ""ProductName"": ""Widget A"", ""Quantity"": 2, ""UnitPrice"": 9.99 },
      { ""ProductName"": ""Widget B"", ""Quantity"": 1, ""UnitPrice"": 19.95 }
    ]
  },
  {
    ""OrderId"": 1002,
    ""CustomerName"": ""Bob Johnson"",
    ""Items"": [
      { ""ProductName"": ""Gadget X"", ""Quantity"": 5, ""UnitPrice"": 4.50 },
      { ""ProductName"": ""Gadget Y"", ""Quantity"": 3, ""UnitPrice"": 7.25 }
    ]
  }
]";
        // Write JSON to a file.
        const string jsonPath = "orders.json";
        File.WriteAllText(jsonPath, json, Encoding.UTF8);

        // Deserialize JSON.
        var orders = JsonConvert.DeserializeObject<List<Order>>(File.ReadAllText(jsonPath, Encoding.UTF8)) ?? new();
        var root = new RootData { Orders = orders };

        // Create the template document.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Begin foreach over orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");

        // Begin foreach over items – the table is created inside this block.
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Start table for the current order's items.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("Product");
        builder.InsertCell(); builder.Writeln("Qty");
        builder.InsertCell(); builder.Writeln("Unit Price");
        builder.InsertCell(); builder.Writeln("Line Total");
        builder.EndRow();

        // Data row for the current item.
        builder.InsertCell(); builder.Writeln("<<[item.ProductName]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.UnitPrice]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity * item.UnitPrice]>>");
        builder.EndRow();

        // End table for this item.
        builder.EndTable();

        // End inner foreach (items).
        builder.Writeln("<</foreach>>");

        // Total amount per order.
        builder.Writeln("Total Amount: <<[order.Items.Sum(i => i.Quantity * i.UnitPrice)]>>");
        builder.Writeln(); // Blank line between orders.

        // End outer foreach (orders).
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, root, "root");

        // Save the final report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
