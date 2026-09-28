using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new()
        {
            Orders = new()
            {
                new Order
                {
                    CustomerName = "Alice",
                    Items = new()
                    {
                        new OrderItem { Name = "Apple", Quantity = 3, UnitPrice = 0.5m },
                        new OrderItem { Name = "Banana", Quantity = 2, UnitPrice = 0.3m }
                    }
                },
                new Order
                {
                    CustomerName = "Bob",
                    Items = new()
                    {
                        new OrderItem { Name = "Orange", Quantity = 5, UnitPrice = 0.4m },
                        new OrderItem { Name = "Grapes", Quantity = 1, UnitPrice = 2.0m }
                    }
                }
            }
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item: <<[item.Name]>> | Qty: <<[item.Quantity]>> | Unit: <<[item.UnitPrice]>> | Total: <<[item.Total]>>");
        builder.Writeln("<</foreach>>");
        // Use the computed OrderTotal property.
        builder.Writeln("Order Total: <<[order.OrderTotal]>>");
        builder.Writeln("<</foreach>>");
        // Grand total across all orders using a pre‑computed property.
        builder.Writeln("Grand Total: <<[model.GrandTotal]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Root model.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();

    // Computed grand total for all orders.
    public decimal GrandTotal => Orders.Sum(o => o.OrderTotal);
}

// Order model with a computed total.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<OrderItem> Items { get; set; } = new();

    // Computed order total for aggregation.
    public decimal OrderTotal => Items.Sum(i => i.Total);
}

// Order item model.
public class OrderItem
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }

    // Computed line total.
    public decimal Total => Quantity * UnitPrice;
}
