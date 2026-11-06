using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Order
{
    public int Id { get; set; } = 0;
    public string CustomerName { get; set; } = "";
    public List<OrderItem> Items { get; set; } = new();
    // Holds JSON representation for debugging.
    public string OrderJson { get; set; } = "";
}

public class OrderItem
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; } = 0;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var order = new Order
        {
            Id = 123,
            CustomerName = "John Doe",
            Items = new()
            {
                new OrderItem { ProductName = "Widget", Quantity = 5 },
                new OrderItem { ProductName = "Gadget", Quantity = 2 }
            }
        };

        // Serialize the order object to JSON for debugging.
        order.OrderJson = JsonConvert.SerializeObject(order, Formatting.Indented);

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Order Debug JSON:");
        // Use the correct LINQ Reporting tag referencing the root object name "order".
        builder.Writeln("<<[order.OrderJson]>>");
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated document.
        string outputPath = "Output.docx";
        reportDoc.Save(outputPath);
    }
}
