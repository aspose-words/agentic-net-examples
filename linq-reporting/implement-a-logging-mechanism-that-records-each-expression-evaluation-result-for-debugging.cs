using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document with LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        Order order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 1.20m },
                new Item { Name = "Banana", Price = 0.80m },
                new Item { Name = "Cherry", Price = 2.50m }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save("Report.docx");

        // Write the evaluation log to a file.
        File.WriteAllLines("log.txt", Logger.Entries);
    }
}

// Simple logger that records evaluation messages.
public static class Logger
{
    private static readonly List<string> _entries = new List<string>();
    public static IReadOnlyList<string> Entries => _entries;

    public static void Log(string message)
    {
        _entries.Add($"{DateTime.Now:O} - {message}");
    }
}

// Data model for the report.
public class Order
{
    private string _customerName = "";
    private List<Item> _items = new();

    public string CustomerName
    {
        get
        {
            Logger.Log($"Order.CustomerName evaluated: {_customerName}");
            return _customerName;
        }
        set => _customerName = value ?? "";
    }

    public List<Item> Items
    {
        get
        {
            Logger.Log("Order.Items accessed");
            return _items;
        }
        set => _items = value ?? new List<Item>();
    }
}

// Data model for each item.
public class Item
{
    private string _name = "";
    private decimal _price;

    public string Name
    {
        get
        {
            Logger.Log($"Item.Name evaluated: {_name}");
            return _name;
        }
        set => _name = value ?? "";
    }

    public decimal Price
    {
        get
        {
            Logger.Log($"Item.Price evaluated: {_price}");
            return _price;
        }
        set => _price = value;
    }
}
