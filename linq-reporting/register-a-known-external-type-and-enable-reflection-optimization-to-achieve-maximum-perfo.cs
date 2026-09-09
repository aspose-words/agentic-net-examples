using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization for maximum performance.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a simple data model.
        Order order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Price = 1.20 },
                new Item { Name = "Banana", Price = 0.80 },
                new Item { Name = "Cherry", Price = 2.50 }
            }
        };

        // Build the template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in order.Items]>>");
        // Each row shows the item name and a formatted price using a static helper method.
        builder.Writeln("Item: <<[item.Name]>>  Price: <<[MyHelper.Format(item.Price)]>>");
        builder.Writeln("<</foreach>>");

        // Register the external type so its static members can be used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(MyHelper));

        // Build the report using the root object name "order".
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Simple data model classes.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}

// External type with a static method used in the template.
public static class MyHelper
{
    // Formats a double value as a currency string.
    public static string Format(double value)
    {
        return value.ToString("C2");
    }
}
