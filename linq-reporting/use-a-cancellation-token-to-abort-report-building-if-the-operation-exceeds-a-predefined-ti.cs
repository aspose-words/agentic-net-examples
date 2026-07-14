using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple tag for the customer name.
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        // Loop over order items.
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare data model with a cancellation token.
        // -----------------------------------------------------------------
        using var cts = new CancellationTokenSource();
        // Cancel after 1 second.
        cts.CancelAfter(TimeSpan.FromSeconds(1));
        CancellationToken token = cts.Token;

        // Build a sample order with several items.
        Order order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item("Apple", token),
                new Item("Banana", token),
                new Item("Cherry", token),
                new Item("Date", token),
                new Item("Elderberry", token)
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report with cancellation support.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        bool reportBuilt = false;

        try
        {
            // BuildReport evaluates the data source lazily.
            // The Item.Name getter checks the cancellation token and throws if cancelled.
            reportBuilt = engine.BuildReport(doc, order, "order");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Report generation was canceled due to timeout.");
        }
        catch (AggregateException ae) when (ae.InnerException is OperationCanceledException)
        {
            // Handles the case where the engine wraps the cancellation exception.
            Console.WriteLine("Report generation was canceled due to timeout.");
        }
        catch (InvalidOperationException ex) when (ex.InnerException is OperationCanceledException)
        {
            // Handles the case where the engine wraps the cancellation exception.
            Console.WriteLine("Report generation was canceled due to timeout.");
        }

        if (reportBuilt)
        {
            doc.Save(reportPath);
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Order
{
    // Non‑nullable properties are initialized to avoid warnings.
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    private readonly string _name;
    private readonly CancellationToken _cancellationToken;

    public Item(string name, CancellationToken cancellationToken)
    {
        _name = name;
        _cancellationToken = cancellationToken;
    }

    // The getter simulates work and checks the cancellation token.
    public string Name
    {
        get
        {
            // Simulate a time‑consuming operation.
            Thread.Sleep(300);
            // Abort if cancellation was requested.
            _cancellationToken.ThrowIfCancellationRequested();
            return _name;
        }
    }
}
