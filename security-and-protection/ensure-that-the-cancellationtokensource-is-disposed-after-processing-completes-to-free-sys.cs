using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main()
    {
        // Create a simple document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Protect the document with a password.
        doc.Protect(ProtectionType.ReadOnly, "secret");

        // Define output paths.
        string protectedPath = Path.Combine(Directory.GetCurrentDirectory(), "Protected.docx");
        string unprotectedPath = Path.Combine(Directory.GetCurrentDirectory(), "Unprotected.docx");

        // Save the protected document.
        doc.Save(protectedPath);

        // Use a CancellationTokenSource for a simulated async operation.
        // The using statement guarantees that the token source is disposed after use.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Perform some asynchronous processing that respects cancellation.
            await ProcessDocumentAsync(protectedPath, cts.Token);
            // The CancellationTokenSource will be disposed here.
        }

        // Verify that the protected file was created.
        if (!File.Exists(protectedPath))
            throw new FileNotFoundException("Protected document was not saved.", protectedPath);

        // Load the protected document, then unprotect it using the correct password.
        Document loadedDoc = new Document(protectedPath);
        loadedDoc.Unprotect("secret");

        // Save the unprotected version.
        loadedDoc.Save(unprotectedPath);

        // Verify that the unprotected file was created.
        if (!File.Exists(unprotectedPath))
            throw new FileNotFoundException("Unprotected document was not saved.", unprotectedPath);
    }

    private static async Task ProcessDocumentAsync(string filePath, CancellationToken token)
    {
        // Simulate asynchronous work that loads the document and reads its text.
        await Task.Run(() =>
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Read the document text (just to simulate some processing).
            string text = doc.GetText();

            // Simulate a short delay.
            Thread.Sleep(500);

            // Respect cancellation requests.
            token.ThrowIfCancellationRequested();
        }, token);
    }
}
