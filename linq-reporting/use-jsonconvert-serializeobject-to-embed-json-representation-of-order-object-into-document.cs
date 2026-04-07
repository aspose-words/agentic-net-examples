using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class OrderItem
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = "";
    public List<OrderItem> Items { get; set; } = new();
}

public class ReportModel
{
    public Order Order { get; set; } = new();

    // Returns a pretty‑printed JSON representation of the Order object.
    public string OrderJson => JsonConvert.SerializeObject(Order, Formatting.Indented);
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Order = new Order
            {
                Id = 1001,
                CustomerName = "John Doe",
                Items = new List<OrderItem>
                {
                    new() { Name = "Widget", Quantity = 3 },
                    new() { Name = "Gadget", Quantity = 5 }
                }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("=== Order Debug Information ===");
        // Insert a LINQ Reporting tag that will be replaced with the JSON string.
        builder.Writeln("<<[model.OrderJson]>>");

        // Build the report using the model as the root data source named "model".
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("Report.docx");
    }
}
