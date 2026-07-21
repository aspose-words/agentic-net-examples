using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a tag that references the root object's property.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Products:");

        // Loop over the collection property of the root object.
        builder.Writeln("<<foreach [p in order.Items]>>");
        builder.Writeln("- <<[p.Name]>> : $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Prepare the data model.
        Order order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Price = 1.20 },
                new Item { Name = "Banana", Price = 0.80 },
                new Item { Name = "Orange", Price = 1.50 }
            }
        };

        // Build the report, specifying the collection name ("order") to target this data set.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model class.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Item class used in the collection.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
