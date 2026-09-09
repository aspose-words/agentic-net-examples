using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Iterate over the collection "Items" of the root object "order".
        builder.Writeln("<<foreach [item in order.Items]>>");
        // Output each item's name and price.
        builder.Writeln("Item: <<[item.Name]>> - Price: <<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // After the loop, display the accumulated total using a property on the root object.
        builder.Writeln("Total: <<[order.Total]>>");

        // Prepare sample data.
        Order sampleOrder = new()
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Price = 1.20m },
                new Item { Name = "Banana", Price = 0.80m },
                new Item { Name = "Cherry", Price = 2.50m }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, sampleOrder, "order");

        // Save the generated document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        template.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Root data model.
public class Order
{
    public List<Item> Items { get; set; } = new();

    // Calculated total price of all items.
    public decimal Total => Items.Sum(i => i.Price);
}

// Item model used inside the foreach loop.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
