using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary directory for the sample document.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsDemo");
        Directory.CreateDirectory(tempDir);

        // Path of the sample document.
        string docPath = Path.Combine(tempDir, "Sample.docx");

        // Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        doc.Save(docPath);

        // Prepare a cancelled token.
        CancellationTokenSource cts = new CancellationTokenSource();
        cts.Cancel(); // Cancel before the load operation.

        try
        {
            // Simulate a load that respects cancellation.
            if (cts.Token.IsCancellationRequested)
                throw new OperationCanceledException(cts.Token);

            // Normal load (won't be reached because token is cancelled).
            Document loadedDoc = new Document(docPath);
        }
        catch (OperationCanceledException)
        {
            // Perform necessary cleanup after a cancelled load.
            if (File.Exists(docPath))
                File.Delete(docPath);

            // Remove the temporary directory if it is empty.
            if (Directory.Exists(tempDir) && Directory.GetFileSystemEntries(tempDir).Length == 0)
                Directory.Delete(tempDir);
        }

        // Ensure the program exits without waiting for user input.
    }
}
