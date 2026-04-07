using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization (default is true, set explicitly for clarity).
        ReportingEngine.UseReflectionOptimization = true;

        // Prepare sample hierarchical data.
        Order sampleOrder = new Order
        {
            Id = 1001,
            Customer = new Customer
            {
                Name = "John Doe",
                Email = "john.doe@example.com"
            },
            Items = new List<Item>
            {
                new Item { Name = "Laptop", Price = 1299.99m },
                new Item { Name = "Mouse", Price = 25.50m },
                new Item { Name = "Keyboard", Price = 45.00m }
            }
        };
        // Compute total.
        sampleOrder.Total = 0;
        foreach (var it in sampleOrder.Items)
            sampleOrder.Total += it.Price;

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register external types that can be used inside the template.
        engine.KnownTypes.Add(typeof(MyHelper));
        engine.KnownTypes.Add(typeof(Math)); // example of another type

        // Build the report using the hierarchical data.
        // The root object name in the template is "order".
        engine.BuildReport(doc, sampleOrder, "order");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Creates a simple Word template with LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.Customer.Name]>>");
        builder.Writeln("Email: <<[order.Customer.Email]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[MyHelper.FormatCurrency(item.Price)]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Total: <<[MyHelper.FormatCurrency(order.Total)]>>");

        // Save the template to disk.
        doc.Save(filePath);
    }
}

// Sample data model classes.
public class Order
{
    public int Id { get; set; }
    public Customer Customer { get; set; } = new Customer();
    public List<Item> Items { get; set; } = new List<Item>();
    public decimal Total { get; set; }
}

public class Customer
{
    public string Name { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

// External helper class whose static members can be invoked from the template.
public static class MyHelper
{
    // Formats a decimal value as currency.
    public static string FormatCurrency(decimal value)
    {
        return string.Format("{0:C}", value);
    }
}
