using System;
using System.IO;
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

        // Prepare save options with a progress callback that will cancel the operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };

        string outputPath = Path.Combine(Environment.CurrentDirectory, "Sample.docx");

        try
        {
            // Attempt to save the document. The callback will throw an OperationCanceledException.
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Capture and log the stack trace for debugging purposes.
            Console.WriteLine("OperationCanceledException was caught.");
            Console.WriteLine("Message: " + ex.Message);
            Console.WriteLine("Stack Trace:");
            Console.WriteLine(ex.StackTrace);
        }
    }

    // Callback that aborts the saving process by throwing an OperationCanceledException.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        public void Notify(DocumentSavingArgs args)
        {
            // Immediately cancel the save operation.
            throw new OperationCanceledException(
                $"EstimatedProgress = {args.EstimatedProgress}; Save operation was cancelled for debugging.");
        }
    }
}
