using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create the data model.
        ReportModel model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    Customer = "Alice",
                    Items = new List<Item>
                    {
                        new Item { Name = "Apple",  Quantity = 3, Price = 0.5m },
                        new Item { Name = "Banana", Quantity = 2, Price = 0.3m }
                    }
                },
                new Order
                {
                    Customer = "Bob",
                    Items = new List<Item>
                    {
                        new Item { Name = "Orange", Quantity = 5, Price = 0.4m },
                        new Item { Name = "Grape",  Quantity = 1, Price = 1.2m }
                    }
                }
            }
        };

        // 2. Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer foreach over Orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.Customer]>>");
        builder.Writeln();

        // Inner foreach over Items of the current order.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>: Qty <<[item.Quantity]>>, Price <<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Order total using a pre‑computed property.
        builder.Writeln("Order Total: <<[order.Total]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Grand total across all orders.
        builder.Writeln("Grand Total: <<[model.GrandTotal]>>");

        // 3. Save the template (optional, shown for completeness).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 4. Load the template (demonstrating the load step required before BuildReport).
        Document doc = new Document(templatePath);

        // 5. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Data model classes – all members are public and non‑nullable to avoid warnings.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();

    // Grand total calculated from all orders.
    public decimal GrandTotal => Orders.Sum(o => o.Total);
}

public class Order
{
    public string Customer { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();

    // Total for this order.
    public decimal Total => Items.Sum(i => i.Price * i.Quantity);
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}
