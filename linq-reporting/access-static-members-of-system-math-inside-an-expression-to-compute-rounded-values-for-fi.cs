using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple LINQ Reporting template.
        // Iterate over the collection "Items" and display each item's name and the rounded amount.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Product: <<[item.Name]>>");
        // Use System.Math static method Round to round the amount to 2 decimal places.
        builder.Writeln("Amount (rounded): <<[Math.Round(item.Amount, 2)]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Laptop", Amount = 1299.987 },
                new Item { Name = "Smartphone", Amount = 799.456 },
                new Item { Name = "Headphones", Amount = 199.999 }
            }
        };

        // Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // Register System.Math so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Data model classes must be public with public properties.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Amount { get; set; }
}
