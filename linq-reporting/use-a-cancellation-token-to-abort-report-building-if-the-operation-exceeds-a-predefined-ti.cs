using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static async Task Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report for <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template back.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        Order order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 3 },
                new Item { Name = "Banana", Quantity = 5 },
                new Item { Name = "Cherry", Quantity = 7 }
            }
        };

        // Set up a cancellation token that triggers after 2 seconds.
        using CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
        ReportingEngine engine = new ReportingEngine();

        bool success = false;
        try
        {
            // Run BuildReport in a separate task so we can apply the timeout.
            Task<bool> buildTask = Task.Run(() => engine.BuildReport(doc, order, "order"), cts.Token);
            success = await buildTask;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Report generation was cancelled due to timeout.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Report generation failed: {ex.Message}");
        }

        if (success)
        {
            string resultPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(resultPath);
            Console.WriteLine($"Report generated successfully: {resultPath}");
        }
    }
}

// Public data model classes.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
