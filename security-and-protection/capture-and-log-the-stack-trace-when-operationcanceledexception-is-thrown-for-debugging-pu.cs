using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");

        // Configure save options with a progress callback that will cancel the operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };

        string outputPath = Path.Combine(Path.GetTempPath(), "CanceledSave.docx");

        try
        {
            // Attempt to save the document; the callback will throw OperationCanceledException.
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Capture and log the stack trace for debugging.
            Console.WriteLine("OperationCanceledException caught.");
            Console.WriteLine("Message: " + ex.Message);
            Console.WriteLine("StackTrace:");
            Console.WriteLine(ex.StackTrace);
        }
    }

    // Callback that cancels the save operation almost immediately.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _start = DateTime.Now;
        private const double MaxDuration = 0.0; // Cancel right away.

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _start).TotalSeconds;
            if (elapsed > MaxDuration)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }
}
