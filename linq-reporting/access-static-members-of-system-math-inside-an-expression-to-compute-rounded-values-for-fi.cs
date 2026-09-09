using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class InvoiceItem
{
    public string Description { get; set; } = string.Empty;
    public decimal Amount { get; set; }
}

public class ReportModel
{
    public List<InvoiceItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<InvoiceItem>
            {
                new InvoiceItem { Description = "Consulting", Amount = 1234.567m },
                new InvoiceItem { Description = "Software License", Amount = 2500.0m },
                new InvoiceItem { Description = "Support", Amount = 199.994m }
            }
        };

        // Create a template document in memory.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Invoice Report");
        builder.Writeln("<<foreach [item in Items]>>");
        // Use System.Math static method to round the amount to 2 decimal places.
        builder.Writeln("Item: <<[item.Description]>> - Amount: $<<[Math.Round(item.Amount, 2)]>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register System.Math so its static members can be used in expressions.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("InvoiceReport.docx");
    }
}
