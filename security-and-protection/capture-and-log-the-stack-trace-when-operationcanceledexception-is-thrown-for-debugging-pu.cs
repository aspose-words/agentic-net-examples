using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Configure save options with a progress callback that will cancel the operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };

        try
        {
            // Attempt to save the document. The callback will throw OperationCanceledException.
            doc.Save("output.docx", saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Capture and log the stack trace for debugging.
            Console.WriteLine("Operation was canceled:");
            Console.WriteLine(ex.Message);
            Console.WriteLine("Stack Trace:");
            Console.WriteLine(ex.StackTrace);
        }
    }

    // Callback that cancels the save operation after a short duration.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime;
        private const double MaxDurationSeconds = 0.01; // Very short to trigger cancellation quickly.

        public SavingProgressCallback()
        {
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; Canceled after {elapsed:F2} seconds.");
        }
    }
}
