using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

public class Order
{
    public string Product { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();

    // Calculates the total amount of all orders.
    public decimal Total => Orders.Sum(o => o.Quantity * o.UnitPrice);
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample JSON data representing an array of orders.
        // -----------------------------------------------------------------
        var sampleOrders = new List<Order>
        {
            new Order { Product = "Apple",  Quantity = 3, UnitPrice = 0.50m },
            new Order { Product = "Banana", Quantity = 5, UnitPrice = 0.30m },
            new Order { Product = "Orange", Quantity = 2, UnitPrice = 0.80m }
        };

        string json = JsonConvert.SerializeObject(sampleOrders, Formatting.Indented);
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "orders.json");
        File.WriteAllText(jsonPath, json);

        // -----------------------------------------------------------------
        // 2. Load the JSON into a strongly‑typed model.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Orders = JsonConvert.DeserializeObject<List<Order>>(File.ReadAllText(jsonPath)) ?? new List<Order>()
        };

        // -----------------------------------------------------------------
        // 3. Build the Word template programmatically.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("------------------------------");

        // Begin foreach over the Orders collection.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Product: <<[order.Product]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        builder.Writeln("Unit Price: $<<[order.UnitPrice]>>");
        // Inline arithmetic expression to calculate line amount.
        builder.Writeln("Line Total: $<<[order.Quantity * order.UnitPrice]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln("------------------------------");
        // Use the wrapper's Total property.
        builder.Writeln("Grand Total: $<<[model.Total]>>");

        // -----------------------------------------------------------------
        // 4. Generate the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the name used in the template tags ("model").
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OrderReport.docx");
        doc.Save(outputPath);
    }
}
