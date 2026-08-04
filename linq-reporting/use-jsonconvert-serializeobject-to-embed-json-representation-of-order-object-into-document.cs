using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public List<OrderItem> Items { get; set; } = new();
    public DateTime OrderDate { get; set; }
}

public class OrderItem
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class ReportModel
{
    public Order Order { get; set; } = new();
    public string OrderJson { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            Id = 1001,
            CustomerName = "John Doe",
            OrderDate = DateTime.Now,
            Items = new List<OrderItem>
            {
                new OrderItem { Name = "Widget A", Quantity = 2, Price = 9.99m },
                new OrderItem { Name = "Widget B", Quantity = 1, Price = 19.99m }
            }
        };

        // Serialize the order object to JSON for debugging.
        var model = new ReportModel
        {
            Order = order,
            OrderJson = JsonConvert.SerializeObject(order, Formatting.Indented)
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("Customer: <<[model.Order.CustomerName]>>");
        builder.Writeln("Order ID: <<[model.Order.Id]>>");
        builder.Writeln("Order Date: <<[model.Order.OrderDate]>>");
        builder.Writeln();
        builder.Writeln("Order JSON (debug):");
        builder.Writeln("<<[model.OrderJson]>>");

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
