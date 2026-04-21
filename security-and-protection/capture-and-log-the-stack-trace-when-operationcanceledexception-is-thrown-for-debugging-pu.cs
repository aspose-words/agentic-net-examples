using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Set up save options with a progress callback that will cancel the operation.
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
            // Capture and log the exception details, including the stack trace.
            Console.WriteLine("OperationCanceledException caught.");
            Console.WriteLine("Message: " + ex.Message);
            Console.WriteLine("Stack Trace:");
            Console.WriteLine(ex.StackTrace);
        }
    }

    // Callback that aborts the save operation after a short duration.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _savingStartedAt = DateTime.Now;
        private const double MaxDurationSeconds = 0.01; // Very short to guarantee cancellation.

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _savingStartedAt).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }
}
