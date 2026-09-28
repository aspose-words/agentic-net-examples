using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Central log for expression evaluations.
    public static List<string> EvaluationLog = new();

    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- Item: <<[item.Name]>> Qty: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        Order order = new Order
        {
            OrderId = 12345,
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 5 },
                new Item { Name = "Banana", Quantity = 3 },
                new Item { Name = "Cherry", Quantity = 12 }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);

        // Write the evaluation log to a file.
        string logPath = Path.Combine(outputDir, "evaluation_log.txt");
        File.WriteAllLines(logPath, EvaluationLog);

        // Optionally, also output the log to console.
        Console.WriteLine("Report generated at: " + Path.GetFullPath(reportPath));
        Console.WriteLine("Evaluation log written to: " + Path.GetFullPath(logPath));
    }
}

// Wrapper class for Order with logging in property getters.
public class Order
{
    private int _orderId;
    private string _customerName = "";
    private List<Item> _items = new();

    public int OrderId
    {
        get
        {
            Program.EvaluationLog.Add($"OrderId evaluated: {_orderId}");
            return _orderId;
        }
        set => _orderId = value;
    }

    public string CustomerName
    {
        get
        {
            Program.EvaluationLog.Add($"CustomerName evaluated: {_customerName}");
            return _customerName;
        }
        set => _customerName = value ?? "";
    }

    public List<Item> Items
    {
        get
        {
            Program.EvaluationLog.Add($"Items collection accessed, count: {_items.Count}");
            return _items;
        }
        set => _items = value ?? new List<Item>();
    }
}

// Wrapper class for Item with logging in property getters.
public class Item
{
    private string _name = "";
    private int _quantity;

    public string Name
    {
        get
        {
            Program.EvaluationLog.Add($"Item.Name evaluated: {_name}");
            return _name;
        }
        set => _name = value ?? "";
    }

    public int Quantity
    {
        get
        {
            Program.EvaluationLog.Add($"Item.Quantity evaluated: {_quantity}");
            return _quantity;
        }
        set => _quantity = value;
    }
}
