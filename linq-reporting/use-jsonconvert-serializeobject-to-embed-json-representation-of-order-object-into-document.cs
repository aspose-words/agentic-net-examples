using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Order
{
    public int Id { get; set; } = 1;
    public string CustomerName { get; set; } = "John Doe";
    public List<OrderItem> Items { get; set; } = new()
    {
        new OrderItem { Product = "Widget", Quantity = 3, Price = 9.99 },
        new OrderItem { Product = "Gadget", Quantity = 2, Price = 14.50 }
    };
    public DateTime OrderDate { get; set; } = DateTime.Now;
}

public class OrderItem
{
    public string Product { get; set; } = "";
    public int Quantity { get; set; } = 0;
    public double Price { get; set; } = 0.0;
}

public class ReportModel
{
    public Order Order { get; set; } = new();

    // Serialized JSON representation of the Order object for debugging.
    public string OrderJson => JsonConvert.SerializeObject(Order, Formatting.Indented);
}

public class Program
{
    public static void Main()
    {
        // Prepare the data model.
        var model = new ReportModel();

        // Create a new blank document that will serve as the template.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a heading and a LINQ Reporting tag that will output the JSON string.
        builder.Writeln("Order JSON Debug:");
        builder.Writeln("<<[model.OrderJson]>>");

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
