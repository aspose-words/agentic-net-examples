using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public List<OrderItem> Items { get; set; } = new();

    // Exclude this property from JSON serialization to avoid recursive calls.
    [JsonIgnore]
    public string Json => JsonConvert.SerializeObject(this, Formatting.Indented);
}

public class OrderItem
{
    public int ItemId { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }

    public OrderItem(int itemId, string name, int quantity)
    {
        ItemId = itemId;
        Name = name;
        Quantity = quantity;
    }
}

public class Program
{
    public static void Main()
    {
        // Required for some encodings used by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var order = new Order
        {
            OrderId = 1001,
            CustomerName = "John Doe",
            Items = new List<OrderItem>
            {
                new OrderItem(1, "Laptop", 2),
                new OrderItem(2, "Mouse", 5)
            }
        };

        // Build the template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Order Debug Information:");
        // LINQ Reporting tag that outputs the JSON string.
        builder.Writeln("<<[order.Json]>>");

        // Generate the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the result.
        const string outputPath = "OrderReport.docx";
        doc.Save(outputPath);
    }
}
