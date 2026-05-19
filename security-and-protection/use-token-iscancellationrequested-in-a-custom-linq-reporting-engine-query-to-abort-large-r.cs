using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple non‑anonymous data source required by ReportingEngine.
    public class DataSource
    {
        public IEnumerable<int> ds { get; set; }
    }

    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Report.docx");

        // Create a template document containing a foreach tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Large Report:");
        // The tag iterates over the data source named 'ds'.
        builder.Writeln("<<foreach [item in ds]>>Item: <<[item]>>\n<</foreach>>");

        // Generate a large data source.
        List<int> numbers = Enumerable.Range(1, 1_000_000).ToList();

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromMilliseconds(10));
        CancellationToken token = cts.Token;

        // Wrap the data source with a LINQ query that checks the cancellation token.
        IEnumerable<int> cancellableNumbers = numbers.Where(i =>
        {
            if (token.IsCancellationRequested)
                throw new OperationCanceledException("Report generation was cancelled.");
            return true;
        });

        // Prepare the data object for the reporting engine.
        var data = new DataSource { ds = cancellableNumbers };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        bool reportBuilt = false;
        try
        {
            // No need to specify a data source name because the property name 'ds' is used in the template.
            reportBuilt = engine.BuildReport(template, data);
        }
        catch (InvalidOperationException ex) when (ex.InnerException is OperationCanceledException)
        {
            // The ReportingEngine wraps the cancellation exception; treat it as a cancellation.
            Console.WriteLine(ex.InnerException.Message);
        }
        catch (OperationCanceledException ex)
        {
            // Fallback in case the engine throws the exception directly.
            Console.WriteLine(ex.Message);
        }

        // Save the document only if the report was built successfully.
        if (reportBuilt)
        {
            template.Save(outputPath);
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The report file was not created.");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Report generation was aborted; no file was saved.");
        }
    }
}
