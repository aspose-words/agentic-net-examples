using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags.
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Prepare the data model with logging in property getters.
        ReportModel model = new ReportModel
        {
            CustomerName = "Acme Corp",
            Items = new List<Item>
            {
                new Item { Name = "Widget", Price = 9.99 },
                new Item { Name = "Gadget", Price = 14.50 }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("ReportOutput.docx");

        // Output the evaluation log.
        Console.WriteLine("Expression Evaluation Log:");
        foreach (string entry in ReportModel.EvaluationLog)
        {
            Console.WriteLine(entry);
        }
    }
}

// Root data model with logging.
public class ReportModel
{
    // Shared log for all evaluations.
    public static List<string> EvaluationLog { get; } = new();

    private string _customerName = string.Empty;
    public string CustomerName
    {
        get
        {
            EvaluationLog.Add($"CustomerName evaluated: {_customerName}");
            return _customerName;
        }
        set => _customerName = value;
    }

    private List<Item> _items = new();
    public List<Item> Items
    {
        get
        {
            EvaluationLog.Add($"Items count evaluated: {_items.Count}");
            return _items;
        }
        set => _items = value;
    }
}

// Item class with logging in each property.
public class Item
{
    private string _name = string.Empty;
    public string Name
    {
        get
        {
            ReportModel.EvaluationLog.Add($"Item.Name evaluated: {_name}");
            return _name;
        }
        set => _name = value;
    }

    private double _price;
    public double Price
    {
        get
        {
            ReportModel.EvaluationLog.Add($"Item.Price evaluated: {_price}");
            return _price;
        }
        set => _price = value;
    }
}
