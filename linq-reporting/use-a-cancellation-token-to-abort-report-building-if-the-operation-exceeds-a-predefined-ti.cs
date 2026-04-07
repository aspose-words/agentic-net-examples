using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    private string _name = string.Empty;

    public ReportModel() { }

    public string Name
    {
        get
        {
            // Simulate a long-running data retrieval operation.
            Thread.Sleep(5000);
            return _name;
        }
        set => _name = value;
    }
}

public class Program
{
    public static async Task Main()
    {
        // Create a simple template document with a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report for <<[model.Name]>>");

        // Prepare the data source.
        ReportModel model = new ReportModel { Name = "Alice" };

        // Set up the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Define a cancellation token that triggers after 2 seconds.
        using CancellationTokenSource cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
        CancellationToken token = cts.Token;

        // Run the report building operation in a separate task.
        Task<bool> buildTask = Task.Run(() => engine.BuildReport(template, model, "model"), token);

        try
        {
            // Wait for either the build to finish or the token to cancel.
            Task completed = await Task.WhenAny(buildTask, Task.Delay(Timeout.Infinite, token));

            if (completed == buildTask && buildTask.IsCompletedSuccessfully)
            {
                // Report built within the time limit – save the result.
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
                template.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            else
            {
                // The operation exceeded the time limit.
                Console.WriteLine("Report generation aborted due to timeout.");
            }
        }
        catch (OperationCanceledException)
        {
            // The cancellation token was triggered.
            Console.WriteLine("Report generation cancelled.");
        }
    }
}
