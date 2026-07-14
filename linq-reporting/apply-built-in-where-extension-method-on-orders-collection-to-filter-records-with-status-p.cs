using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = "";
    public string Status { get; set; } = "";
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template: iterate over the Orders collection and output fields.
        builder.Writeln("<<foreach [order in Model.Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>, Customer: <<[order.CustomerName]>>, Status: <<[order.Status]>>");
        builder.Writeln("<</foreach>>");

        // Sample data.
        List<Order> allOrders = new()
        {
            new Order { Id = 1, CustomerName = "Alice",   Status = "Pending" },
            new Order { Id = 2, CustomerName = "Bob",     Status = "Shipped" },
            new Order { Id = 3, CustomerName = "Charlie", Status = "Pending" },
            new Order { Id = 4, CustomerName = "Diana",   Status = "Cancelled" }
        };

        // Apply the built‑in Where extension method to keep only pending orders.
        List<Order> pendingOrders = allOrders
            .Where(o => o.Status == "Pending")
            .ToList();

        // Wrap the filtered collection in a model object.
        ReportModel model = new ReportModel { Orders = pendingOrders };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "Model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
