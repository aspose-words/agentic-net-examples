using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Sample data model.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Create and save the template document.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Prepare the reporting engine.
        var engine = new ReportingEngine();

        // Cancellation token that triggers after a predefined time limit.
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
        CancellationToken token = cts.Token;

        // Run the report building in a separate task.
        Task<bool> buildTask = Task.Run(() => engine.BuildReport(doc, order, "order"));

        // A task that completes when the cancellation token is triggered.
        Task delayTask = Task.Delay(Timeout.Infinite, token);

        // Wait for either the report to finish or the timeout to occur.
        Task completed = Task.WhenAny(buildTask, delayTask).Result;

        if (completed == buildTask && buildTask.IsCompletedSuccessfully)
        {
            // Report built within the time limit – save the result.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report generated: {outputPath}");
        }
        else
        {
            // Timeout reached – abort the operation.
            Console.WriteLine("Report generation was cancelled due to timeout.");
        }
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln(" - <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

// Public data model classes.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
