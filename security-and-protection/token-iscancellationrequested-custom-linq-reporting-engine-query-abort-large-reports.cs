using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineCancellationDemo
{
    // Simple data model used in the report.
    public class ReportItem
    {
        public int Id { get; set; }
        public string Description { get; set; }
    }

    // Wrapper that aborts enumeration when the supplied CancellationToken is cancelled.
    public class CancellableEnumerable<T> : IEnumerable<T>
    {
        private readonly IEnumerable<T> _source;
        private readonly CancellationToken _token;

        public CancellableEnumerable(IEnumerable<T> source, CancellationToken token)
        {
            _source = source;
            _token = token;
        }

        public IEnumerator<T> GetEnumerator()
        {
            foreach (var item in _source)
            {
                // Abort the report generation as soon as cancellation is requested.
                if (_token.IsCancellationRequested)
                    throw new OperationCanceledException("Report generation was cancelled.", _token);

                yield return item;
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    class Program
    {
        static void Main()
        {
            // Prepare a cancellation token that will be triggered after a short delay.
            var cts = new CancellationTokenSource();
            // For demonstration we cancel after 100 ms.
            cts.CancelAfter(100);
            CancellationToken token = cts.Token;

            // Build a large data set to simulate a heavy report.
            var allItems = Enumerable.Range(1, 100_000)
                                     .Select(i => new ReportItem { Id = i, Description = $"Item {i}" })
                                     .ToList();

            // Wrap the data source with the cancellable enumerable.
            var cancellableData = new CancellableEnumerable<ReportItem>(allItems, token);

            // Create a simple template document in memory.
            // The template uses the LINQ Reporting Engine syntax: <<foreach [in data]>><<[Id]>> - <<[Description]>>\n<</foreach>>
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Report Start");
            builder.Writeln("<<foreach [in data]>>");
            builder.Writeln("Id: <<[Id]>>, Description: <<[Description]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("Report End");

            // Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Ensure that missing members do not cause exceptions.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            try
            {
                // Build the report. The data source name "data" matches the template tag.
                engine.BuildReport(doc, cancellableData, "data");
                // If the operation completes, save the result.
                doc.Save("ReportResult.docx");
                Console.WriteLine("Report generated successfully.");
            }
            catch (InvalidOperationException ex) when (ex.InnerException is OperationCanceledException)
            {
                // Handle cancellation that was wrapped by the ReportingEngine.
                Console.WriteLine($"Report generation aborted: {ex.InnerException.Message}");
            }
            catch (OperationCanceledException ex)
            {
                // Handle direct cancellation (unlikely due to wrapping).
                Console.WriteLine($"Report generation aborted: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors.
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
