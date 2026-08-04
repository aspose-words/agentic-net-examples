using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the demo files.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsDemo_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        string docPath = Path.Combine(tempDir, "sample.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it to the temporary location.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 2. Set up LoadOptions with a progress callback that cancels loading.
        // ---------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new CancelLoadingCallback()
        };

        try
        {
            // Attempt to load the document. The callback will throw
            // OperationCanceledException, causing the load to abort.
            Document loadedDoc = new Document(docPath, loadOptions);
            // If, for any reason, loading succeeds, output the document text.
            Console.WriteLine("Document loaded successfully: " + loadedDoc.GetText().Trim());
        }
        catch (OperationCanceledException ex)
        {
            // -----------------------------------------------------------
            // 3. Handle the cancellation and perform any necessary cleanup.
            // -----------------------------------------------------------
            Console.WriteLine("Loading was cancelled: " + ex.Message);
        }
        finally
        {
            // -----------------------------------------------------------
            // 4. Clean up temporary files and directories.
            // -----------------------------------------------------------
            try
            {
                if (File.Exists(docPath))
                    File.Delete(docPath);
                if (Directory.Exists(tempDir))
                    Directory.Delete(tempDir, true);
            }
            catch
            {
                // Ignored – cleanup should not throw.
            }
        }
    }

    // Callback implementation that unconditionally cancels the load operation.
    private class CancelLoadingCallback : IDocumentLoadingCallback
    {
        public void Notify(DocumentLoadingArgs args)
        {
            throw new OperationCanceledException(
                $"Loading cancelled at {args.EstimatedProgress}% progress.");
        }
    }
}
