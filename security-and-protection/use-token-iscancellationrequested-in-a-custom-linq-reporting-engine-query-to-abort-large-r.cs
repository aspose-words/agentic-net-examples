using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple source document.
        string sourcePath = "source.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document for protection.");
        doc.Save(sourcePath);

        // Protect the document with a password.
        string protectedPath = "protected.docx";
        doc.Protect(ProtectionType.ReadOnly, "secret");
        doc.Save(protectedPath);

        // Simulate large report generation with cancellation support.
        var cts = new CancellationTokenSource();
        // Cancel after a short delay to demonstrate aborting.
        cts.CancelAfter(10); // milliseconds

        try
        {
            string report = GenerateReport(cts.Token);
            // Save the report to a file if it completes.
            File.WriteAllText("report.txt", report);
            Console.WriteLine("Report generated successfully.");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Report generation was cancelled.");
        }

        // Verify that the protected document was saved.
        if (!File.Exists(protectedPath))
            throw new Exception("Protected document was not saved.");

        Console.WriteLine("Finished.");
    }

    private static string GenerateReport(CancellationToken token)
    {
        // Large data set to simulate a heavy report.
        var data = Enumerable.Range(1, 1_000_000);

        // LINQ query that checks for cancellation.
        var sb = data.Select(i =>
            {
                if (token.IsCancellationRequested)
                    throw new OperationCanceledException(token);
                // Simulate some processing work.
                return i * i;
            })
            // Take a subset to keep the output manageable if not cancelled.
            .Take(100)
            .Aggregate(new StringBuilder(), (builder, value) => builder.AppendLine(value.ToString()));

        return sb.ToString();
    }
}
