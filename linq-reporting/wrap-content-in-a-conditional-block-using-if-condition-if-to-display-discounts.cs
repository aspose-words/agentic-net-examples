using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to construct the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write static text and a placeholder for the customer's name.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin a foreach loop over the collection of items.
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Write item details.
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("Price: $<<[item.Price]>>");

        // Conditional block: display discount only when it is greater than zero.
        builder.Writeln("<<if [item.Discount > 0]>>Discount: $<<[item.Discount]>> <</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Build the report using a sample data model.
        Order sampleOrder = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Laptop", Price = 1200.00m, Discount = 100.00m },
                new Item { Name = "Mouse", Price = 25.00m, Discount = 0.00m },
                new Item { Name = "Keyboard", Price = 45.00m, Discount = 5.00m }
            }
        };

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, sampleOrder, "order");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model representing an order.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Item model used inside the order.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public decimal Discount { get; set; }
}
