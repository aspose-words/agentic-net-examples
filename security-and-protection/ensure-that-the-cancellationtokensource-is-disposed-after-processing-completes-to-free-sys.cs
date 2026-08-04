using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        const string outputPath = "ProtectedDocument.docx";

        // Create a blank document, add some text and protect it with a password.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sensitive information that should be read‑only.");
        doc.Protect(ProtectionType.ReadOnly, "SecretPassword");
        doc.Save(outputPath);

        // Validate that the document was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Perform a sample processing operation that uses a CancellationTokenSource.
        ProcessDocument(outputPath);
    }

    private static void ProcessDocument(string path)
    {
        // The CancellationTokenSource is wrapped in a using statement to guarantee disposal.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Simulate an asynchronous operation that respects cancellation.
            Task processingTask = Task.Run(() =>
            {
                // Load the protected document (no password needed for programmatic access).
                Document loadedDoc = new Document(path);
                // Perform a trivial operation – retrieve the document text length.
                string text = loadedDoc.GetText();
                Console.WriteLine($"Loaded document text length: {text.Length}");
            }, cts.Token);

            // Wait for the task to complete.
            processingTask.Wait();
        } // cts is disposed here, freeing system resources.
    }
}
