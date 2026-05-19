using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple blank document and add some text.
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
            // Capture and log the exception message and stack trace for debugging.
            Console.WriteLine("Caught OperationCanceledException: " + ex.Message);
            Console.WriteLine("Stack Trace:");
            Console.WriteLine(ex.StackTrace);
        }
    }

    // Callback that aborts the saving process by throwing OperationCanceledException.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        public void Notify(DocumentSavingArgs args)
        {
            // Immediately cancel the operation.
            throw new OperationCanceledException($"Saving canceled at progress {args.EstimatedProgress}%.");
        }
    }
}
