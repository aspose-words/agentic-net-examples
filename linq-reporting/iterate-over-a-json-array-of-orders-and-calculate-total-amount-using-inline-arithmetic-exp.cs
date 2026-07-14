using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting; // ReportingEngine namespace
using Aspose.Words.Reporting; // Ensure ReportingEngine is available
using Aspose.Words.Reporting; // For JsonDataSource
using Aspose.Words.Reporting; // For JsonDataLoadOptions
using Newtonsoft.Json;

public class Order
{
    public string CustomerName { get; set; } = "";
    public double Amount { get; set; }
}

public class OrdersRoot
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "orders.json";
        var sampleData = new OrdersRoot
        {
            Orders = new List<Order>
            {
                new Order { CustomerName = "Alice", Amount = 120.5 },
                new Order { CustomerName = "Bob", Amount = 80.0 },
                new Order { CustomerName = "Charlie", Amount = 45.75 }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Configure JSON data load options to always generate a root object.
        var jsonLoadOptions = new Aspose.Words.Reporting.JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };

        // Load JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, jsonLoadOptions);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Orders Report");
        // Loop through each order.
        builder.Writeln("<<foreach [order in data.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>> - Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");
        // Calculate total using an inline arithmetic expression.
        builder.Writeln("Total Amount: <<[data.Orders.Sum(o => o.Amount)]>>");

        // Build the report using the JSON data source.
        ReportingEngine engine = new ReportingEngine();
        // Pass the data source name "data" so that template tags can reference it.
        engine.BuildReport(template, jsonDataSource, "data");

        // Save the generated report.
        template.Save("OrdersReport.docx");
    }
}
