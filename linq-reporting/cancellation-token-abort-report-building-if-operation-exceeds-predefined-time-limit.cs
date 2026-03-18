using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportCancellation
{
    // Simple data source used by the report.
    public class SampleData
    {
        public string Title { get; set; } = "Report Title";
        public int Value { get; set; } = 42;
    }

    class Program
    {
        // Maximum allowed time for the report building operation (in seconds).
        private const double MaxDurationSeconds = 2.0;

        static void Main()
        {
            // Create a cancellation token source that will be cancelled after the timeout.
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(MaxDurationSeconds));
            CancellationToken token = cts.Token;

            try
            {
                // Load the template document.
                // (If the template is large, you could also attach a loading progress callback that checks the token.)
                Document template = new Document("Template.docx");

                // Prepare the reporting engine.
                ReportingEngine engine = new ReportingEngine();

                // Prepare the data source.
                SampleData data = new SampleData();

                // Run the report building in a separate task so we can enforce the timeout.
                Task<bool> buildTask = Task.Run(() =>
                {
                    // Periodically check the cancellation token.
                    if (token.IsCancellationRequested)
                        throw new OperationCanceledException(token);

                    // Build the report.
                    bool success = engine.BuildReport(template, data, "ds");

                    // Check again after building.
                    token.ThrowIfCancellationRequested();

                    return success;
                }, token);

                // Wait for the task to complete within the allowed time.
                if (!buildTask.Wait(TimeSpan.FromSeconds(MaxDurationSeconds)))
                {
                    // Timeout elapsed – cancel the operation.
                    cts.Cancel();
                    throw new OperationCanceledException("Report building exceeded the time limit.", token);
                }

                // If we reach this point, the report was built successfully within the time limit.
                bool result = buildTask.Result;
                Console.WriteLine($"Report built successfully: {result}");

                // Save the generated report.
                template.Save("ReportResult.docx");
            }
            catch (OperationCanceledException ex)
            {
                // Handle the cancellation (e.g., log, inform the user, clean up).
                Console.WriteLine($"Operation canceled: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle other possible exceptions.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
