using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class Logger
{
    private static readonly List<string> _logs = new();

    public static void Log(string message) => _logs.Add(message);

    public static IEnumerable<string> Logs => _logs;

    public static void Clear() => _logs.Clear();
}

public class Item
{
    private int _index;
    private string _name = string.Empty;

    public int Index
    {
        get
        {
            Logger.Log($"Item.Index evaluated: {_index}");
            return _index;
        }
        set => _index = value;
    }

    public string Name
    {
        get
        {
            Logger.Log($"Item.Name evaluated: {_name}");
            return _name;
        }
        set => _name = value ?? string.Empty;
    }
}

public class Order
{
    private string _customerName = string.Empty;
    private List<Item> _items = new();

    public string CustomerName
    {
        get
        {
            Logger.Log($"Order.CustomerName evaluated: {_customerName}");
            return _customerName;
        }
        set => _customerName = value ?? string.Empty;
    }

    public List<Item> Items
    {
        get
        {
            Logger.Log($"Order.Items evaluated: count {_items.Count}");
            return _items;
        }
        set => _items = value ?? new();
    }

    public Order()
    {
        CustomerName = "Acme Corp";
        Items = new()
        {
            new Item { Index = 1, Name = "Widget" },
            new Item { Index = 2, Name = "Gadget" },
            new Item { Index = 3, Name = "Doohickey" }
        };
    }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var order = new Order();

        // Create template document.
        var templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln(" - Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load template and build report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        bool success = engine.BuildReport(doc, order, "order");

        // Save the generated report.
        var outputPath = "report.docx";
        doc.Save(outputPath);

        // Output logging information.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine("Expression evaluation log:");
        foreach (var log in Logger.Logs)
        {
            Console.WriteLine(log);
        }

        // Optionally write log to a file.
        File.WriteAllLines("log.txt", Logger.Logs);
    }
}
