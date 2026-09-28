using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Sample Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 2. Load the template for report generation.
        Document doc = new Document(templatePath);

        // 3. Prepare sample data.
        ReportModel model = new ReportModel();
        for (int i = 1; i <= 5000; i++)
        {
            model.Items.Add(new Item { Index = i, Name = $"Product {i}" });
        }

        // 4. Build the report with a cancellation token that aborts after a time limit.
        ReportingEngine engine = new ReportingEngine();

        using CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
        Task<bool> buildTask = Task.Run(() =>
        {
            // The BuildReport method returns a bool indicating success when InlineErrorMessages is used.
            // Here we just call it; the result is not needed for this example.
            engine.BuildReport(doc, model, "model");
            return true;
        }, cts.Token);

        try
        {
            // Wait for the task to complete within the timeout.
            bool completed = buildTask.Wait(TimeSpan.FromSeconds(2), cts.Token);
            if (completed)
            {
                // Report built successfully within the time limit.
                string resultPath = Path.Combine(outputDir, "ReportOutput.docx");
                doc.Save(resultPath);
                Console.WriteLine($"Report generated: {resultPath}");
            }
            else
            {
                // Timeout occurred; report generation is considered aborted.
                Console.WriteLine("Report generation aborted due to timeout.");
            }
        }
        catch (OperationCanceledException)
        {
            // The cancellation token triggered cancellation.
            Console.WriteLine("Report generation cancelled.");
        }
        catch (AggregateException ae) when (ae.InnerException is OperationCanceledException)
        {
            Console.WriteLine("Report generation cancelled via aggregate exception.");
        }
    }
}
