using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    // Reusable helper that processes a document safely, respecting cancellation.
    public static async Task SafeProcessDocumentAsync(string sourcePath, string destinationPath, CancellationToken cancellationToken)
    {
        // Throw if cancellation was requested before we start.
        cancellationToken.ThrowIfCancellationRequested();

        // Load the document. No password is needed for this sample.
        Document doc = new Document(sourcePath, new LoadOptions());

        // Check cancellation again after a potentially time‑consuming operation.
        cancellationToken.ThrowIfCancellationRequested();

        // Apply write protection with a password and recommend read‑only opening.
        doc.WriteProtection.SetPassword("SecretPwd");
        doc.WriteProtection.ReadOnlyRecommended = true;

        // Prepare save options that encrypt the document with the same password.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "SecretPwd"
        };

        // Save the protected document.
        doc.Save(destinationPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(destinationPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Final cancellation check before completing.
        cancellationToken.ThrowIfCancellationRequested();

        await Task.CompletedTask; // Placeholder for async compatibility.
    }

    public static async Task Main()
    {
        // Prepare a temporary folder for the demo files.
        string demoFolder = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(demoFolder);

        string sourceFile = Path.Combine(demoFolder, "Source.docx");
        string outputFile = Path.Combine(demoFolder, "Protected.docx");

        // Create a simple source document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello Aspose.Words! This document will be protected.");
        sourceDoc.Save(sourceFile);

        // Set up a cancellation token (no cancellation in this example).
        using CancellationTokenSource cts = new CancellationTokenSource();

        try
        {
            await SafeProcessDocumentAsync(sourceFile, outputFile, cts.Token);
            // Indicate success (no interactive input required).
            Console.WriteLine("Document processed and saved to: " + outputFile);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document processing was canceled.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
