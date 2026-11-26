using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempFolder);

        // Create a simple document and save it.
        string docPath = Path.Combine(tempFolder, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        doc.Save(docPath);

        // Set up a cancelled token.
        CancellationTokenSource cts = new CancellationTokenSource();
        cts.Cancel(); // Simulate cancellation before loading.

        try
        {
            // Manually check for cancellation before loading the document.
            if (cts.Token.IsCancellationRequested)
                throw new OperationCanceledException(cts.Token);

            // Load the document (no special load options needed for this demo).
            Document loadedDoc = new Document(docPath);
            Console.WriteLine("Document loaded successfully (unexpected).");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document loading was cancelled as expected.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected exception: {ex.GetType().Name} - {ex.Message}");
        }
        finally
        {
            // Clean up temporary files.
            if (File.Exists(docPath))
                File.Delete(docPath);
            if (Directory.Exists(tempFolder))
                Directory.Delete(tempFolder, true);
        }
    }
}
