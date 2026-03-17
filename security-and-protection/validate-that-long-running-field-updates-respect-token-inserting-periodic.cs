using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

namespace FieldUpdateCancellationDemo
{
    // Implements progress callback that checks a CancellationToken on each progress notification.
    public class CancellationFieldProgressCallback : IFieldUpdatingProgressCallback
    {
        private readonly CancellationToken _cancellationToken;

        public CancellationFieldProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        // This method is called periodically during a long‑running field update.
        // If cancellation has been requested, abort the operation by throwing.
        public void Notify(FieldUpdatingProgressArgs args)
        {
            // Optional: you can also inspect args.UpdateCompleted, args.TotalFieldsCount, etc.
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException("Field updating was canceled by the user.");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a cancellation token source that could be triggered from elsewhere.
            var cts = new CancellationTokenSource();

            // For demonstration, cancel after a short delay (e.g., 100 ms).
            // In real scenarios, cancellation would be driven by UI or other logic.
            var timer = new System.Timers.Timer(100) { AutoReset = false };
            timer.Elapsed += (s, e) => cts.Cancel();
            timer.Start();

            // Load or create a document.
            Document doc = new Document(); // Empty document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a large number of fields to simulate a long‑running update.
            for (int i = 0; i < 5000; i++)
            {
                // InsertField overload takes the field code string (and optional field value).
                builder.InsertField($"= {i} + {i + 1}");
                builder.Writeln(); // Add line break for readability.
            }

            // Assign the progress callback that checks the cancellation token.
            doc.FieldOptions.FieldUpdatingProgressCallback = new CancellationFieldProgressCallback(cts.Token);

            try
            {
                // This will invoke the callback periodically.
                doc.UpdateFields();
                Console.WriteLine("Field update completed successfully.");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Field update canceled: {ex.Message}");
            }

            // Save the document (optional, demonstrates normal save flow).
            // The save operation itself can also have its own progress callback if needed.
            doc.Save("Result.docx", SaveFormat.Docx);
        }
    }
}
