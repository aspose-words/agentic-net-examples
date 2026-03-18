using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FieldUpdateCancellationDemo
{
    // Implements progress callback that checks the cancellation token on each progress notification.
    public class FieldUpdatingProgressCallback : IFieldUpdatingProgressCallback
    {
        private readonly CancellationToken _cancellationToken;

        public FieldUpdatingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // This method is called by Aspose.Words during field updating.
        // ThrowIfCancellationRequested aborts the operation if cancellation was requested.
        void IFieldUpdatingProgressCallback.Notify(FieldUpdatingProgressArgs args)
        {
            // Abort the field update loop as soon as cancellation is signaled.
            _cancellationToken.ThrowIfCancellationRequested();

            // Optional: you can log progress here.
            Console.WriteLine($"Updated {args.UpdatedFieldsCount}/{args.TotalFieldsCount} fields. Completed: {args.UpdateCompleted}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a cancellation token source that will cancel after a short delay.
            using var cts = new CancellationTokenSource();
            // For demonstration, cancel after 500 ms.
            cts.CancelAfter(500);
            CancellationToken token = cts.Token;

            // Load or create a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a number of fields to simulate a long‑running update.
            for (int i = 0; i < 1000; i++)
            {
                // Build the field code manually – InsertField(string) is the overload available in all supported versions.
                builder.InsertField($"= {i} * {i + 1}");
                builder.Writeln(); // separate fields with line breaks
            }

            // Assign the progress callback that checks the cancellation token.
            doc.FieldOptions.FieldUpdatingProgressCallback = new FieldUpdatingProgressCallback(token);

            try
            {
                // This will invoke the callback for each field update.
                doc.UpdateFields();
                Console.WriteLine("All fields updated successfully.");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Field updating was canceled.");
            }

            // Optionally save the partially updated document.
            // Note: Save operation is not required for the cancellation demo.
            // doc.Save("PartialResult.docx");
        }
    }
}
