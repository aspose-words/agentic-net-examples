using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Order
{
    public int OrderId { get; set; } = 0;
    public string CustomerName { get; set; } = "";
    public List<OrderItem> Items { get; set; } = new();
}

public class OrderItem
{
    public string Product { get; set; } = "";
    public int Quantity { get; set; } = 0;
    public decimal Price { get; set; } = 0;
}

public class ReportModel
{
    public Order Order { get; set; } = new();
    public string JsonDebug { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Sample order data
        var order = new Order
        {
            OrderId = 123,
            CustomerName = "John Doe",
            Items = new List<OrderItem>
            {
                new OrderItem { Product = "Apple", Quantity = 3, Price = 0.5m },
                new OrderItem { Product = "Banana", Quantity = 5, Price = 0.3m }
            }
        };

        // Serialize order to JSON for debugging
        string json = JsonConvert.SerializeObject(order, Formatting.Indented);

        // Wrap data in a model for LINQ Reporting
        var model = new ReportModel
        {
            Order = order,
            JsonDebug = json
        };

        // Create template document with a placeholder for JSON
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Order JSON Debug:");
        builder.Writeln("<<[model.JsonDebug]>>");
        templateDoc.Save(templatePath);

        // Load template and build the report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
