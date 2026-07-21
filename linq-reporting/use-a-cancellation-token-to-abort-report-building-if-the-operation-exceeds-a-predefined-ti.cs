using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header.
        builder.Writeln("Customer: <<[model.Order.CustomerName]>>");
        builder.Writeln();

        // Repeating section for order items.
        builder.Writeln("<<foreach [item in model.Order.Items]>>");
        builder.Writeln("Item: <<[item.Name]>> - $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare data model with a cancellation token.
        // -----------------------------------------------------------------
        // Cancel the operation if it runs longer than 500 milliseconds.
        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(500));
        CancellationToken token = cts.Token;

        ReportModel model = new()
        {
            Order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item("Widget", 19.99, token),
                    new Item("Gadget", 29.99, token),
                    new Item("Doohickey", 9.99, token)
                }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report with cancellation support.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            // Inline error messages will be inserted instead of throwing for missing members.
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // BuildReport returns a bool indicating whether parsing succeeded.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // If the token was cancelled, treat the operation as aborted.
        if (token.IsCancellationRequested)
        {
            Console.WriteLine("Report generation was canceled due to timeout.");
            return;
        }

        if (success)
        {
            reportDoc.Save("Report.docx");
            Console.WriteLine("Report generated successfully.");
        }
        else
        {
            Console.WriteLine("Report generation failed due to template errors.");
        }
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public ReportModel() { }

    public CancellationToken Token { get; set; }

    public Order Order { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    private readonly string _name;
    private readonly CancellationToken _token;

    public Item(string name, double price, CancellationToken token)
    {
        _name = name;
        Price = price;
        _token = token;
    }

    public string Name
    {
        get
        {
            // Simulate a time‑consuming operation.
            Thread.Sleep(200);

            // If cancellation was requested, return a placeholder instead of throwing.
            if (_token.IsCancellationRequested)
                return "[canceled]";

            return _name;
        }
    }

    public double Price { get; }
}
