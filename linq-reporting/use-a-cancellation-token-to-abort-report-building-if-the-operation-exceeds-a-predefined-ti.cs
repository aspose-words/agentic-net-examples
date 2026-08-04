using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // 1. Create a LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a foreach tag that will iterate over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template back (required before building the report).
        Document loadedTemplate = new Document(templatePath);

        // 3. Prepare a large data source to make the build take noticeable time.
        ReportModel model = new ReportModel();
        for (int i = 0; i < 200_000; i++)
        {
            model.Items.Add(new Item { Name = $"Item #{i + 1}" });
        }

        // 4. Set up a cancellation token that will trigger after a predefined timeout.
        using CancellationTokenSource cts = new(TimeSpan.FromSeconds(2));
        CancellationToken token = cts.Token;

        // 5. Run the report building in a separate task so we can monitor the timeout.
        Task<bool> buildTask = Task.Run(() =>
        {
            // Check for cancellation before starting the heavy operation.
            token.ThrowIfCancellationRequested();

            ReportingEngine engine = new ReportingEngine();
            // BuildReport returns a bool only when InlineErrorMessages is set; we ignore the return value here.
            engine.BuildReport(loadedTemplate, model, "model");
            return true;
        }, token);

        try
        {
            // Wait for either the build to finish or the timeout to elapse.
            bool completed = buildTask.Wait(TimeSpan.FromSeconds(5), token);
            if (completed && buildTask.IsCompletedSuccessfully)
            {
                // Save the generated report if the build finished in time.
                loadedTemplate.Save(reportPath);
                Console.WriteLine($"Report generated successfully: {reportPath}");
            }
            else
            {
                Console.WriteLine("Report building was cancelled due to timeout.");
            }
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Report building was cancelled via cancellation token.");
        }
        finally
        {
            // Clean up temporary files (optional).
            if (File.Exists(templatePath))
                File.Delete(templatePath);
        }
    }
}
