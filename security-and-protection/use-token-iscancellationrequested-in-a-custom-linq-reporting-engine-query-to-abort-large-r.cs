using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LargeDataSource : IEnumerable<int>
{
    private readonly int _count;
    private readonly CancellationToken _token;

    public LargeDataSource(int count, CancellationToken token)
    {
        _count = count;
        _token = token;
    }

    public IEnumerator<int> GetEnumerator()
    {
        for (int i = 0; i < _count; i++)
        {
            // Abort if cancellation is requested.
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException("Report generation was canceled.");

            yield return i;
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

public class Program
{
    public static void Main()
    {
        // Create a simple template document with a LINQ Reporting Engine tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report:");
        // Correct foreach syntax for the Reporting Engine.
        builder.Writeln("<<foreach [item in Data]>><<[item]>><</foreach>>");

        // Set up a cancellation token that will trigger after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            cts.CancelAfter(10); // milliseconds

            // Create a large data source that checks the token during enumeration.
            LargeDataSource data = new LargeDataSource(1_000_000, cts.Token);

            ReportingEngine engine = new ReportingEngine();

            try
            {
                // Build the report using the data source named "Data".
                engine.BuildReport(template, data, "Data");

                // If the report completes, save the full document.
                template.Save("Report.docx");
                Console.WriteLine("Report generated successfully.");
            }
            // The ReportingEngine wraps the OperationCanceledException in an InvalidOperationException.
            catch (InvalidOperationException ex) when (ex.InnerException is OperationCanceledException)
            {
                // On cancellation, save the partially generated document (if any).
                template.Save("ReportPartial.docx");
                Console.WriteLine($"Report generation aborted: {ex.InnerException.Message}");
            }
            // Fallback in case a raw OperationCanceledException is thrown.
            catch (OperationCanceledException ex)
            {
                template.Save("ReportPartial.docx");
                Console.WriteLine($"Report generation aborted: {ex.Message}");
            }
        }
    }
}
