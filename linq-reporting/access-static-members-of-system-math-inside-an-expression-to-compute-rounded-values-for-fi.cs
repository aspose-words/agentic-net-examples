using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public decimal Amount { get; set; } = 0m;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Amount = 123.4567m },
                new Item { Amount = 89.1234m },
                new Item { Amount = 45.9876m }
            }
        };

        // Create a blank document and build the template.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        // Write the original amount.
        builder.Writeln("Original: <<[item.Amount]>>");
        // Write the amount rounded to 2 decimal places using System.Math.Round.
        builder.Writeln("Rounded: <<[Math.Round(item.Amount, 2)]>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Allow the template to use static members of System.Math.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("RoundedReport.docx");
    }
}
