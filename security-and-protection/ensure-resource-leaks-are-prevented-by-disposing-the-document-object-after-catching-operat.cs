using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = "Output.docx";

        // Create a new blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Set up save options with a progress callback that cancels the operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new CancelSavingCallback()
        };

        try
        {
            // Attempt to save the document. The callback will throw OperationCanceledException.
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Handle the cancellation.
            Console.WriteLine("Document saving was canceled: " + ex.Message);
        }
        finally
        {
            // Dispose the Document if it implements IDisposable to prevent resource leaks.
            if (doc is IDisposable disposable)
                disposable.Dispose();
        }

        // Verify that the file was not created due to cancellation.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("File was saved unexpectedly.");
        }
        else
        {
            Console.WriteLine("No output file was created, as expected.");
        }
    }

    // Callback that cancels the save operation immediately.
    private class CancelSavingCallback : IDocumentSavingCallback
    {
        public void Notify(DocumentSavingArgs args)
        {
            throw new OperationCanceledException(
                $"EstimatedProgress = {args.EstimatedProgress}; Canceled at {DateTime.Now}");
        }
    }
}
