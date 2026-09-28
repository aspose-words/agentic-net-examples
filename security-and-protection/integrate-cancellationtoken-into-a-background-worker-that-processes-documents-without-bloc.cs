using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;      // Needed for LoadOptions
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main()
    {
        // Paths for source and output documents.
        string sourcePath = "source.docx";
        string protectedPath = "protected.docx";

        // Create a simple source document.
        Document sourceDoc = new Document();
        Paragraph para = new Paragraph(sourceDoc);
        para.AppendChild(new Run(sourceDoc, "Sample text for protection demo."));
        sourceDoc.FirstSection.Body.AppendChild(para);
        sourceDoc.Save(sourcePath);

        // Set up cancellation.
        using var cts = new CancellationTokenSource();

        // Start background processing.
        Task processingTask = ProcessDocumentAsync(sourcePath, protectedPath, cts.Token);

        // Simulate a short wait before cancelling.
        await Task.Delay(500);
        cts.Cancel();

        try
        {
            await processingTask;
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document processing was cancelled.");
        }

        // Verify result.
        if (File.Exists(protectedPath))
        {
            // Load the saved document to confirm it is protected.
            LoadOptions loadOptions = new LoadOptions { Password = "pwd123" };
            Document protectedDoc = new Document(protectedPath, loadOptions);
            if (protectedDoc.ProtectionType != ProtectionType.NoProtection)
                Console.WriteLine("Protected document saved successfully.");
            else
                throw new InvalidOperationException("Document was saved but not protected.");
        }
        else
        {
            Console.WriteLine("Protected document was not created due to cancellation.");
        }

        // Clean up sample files.
        if (File.Exists(sourcePath)) File.Delete(sourcePath);
        if (File.Exists(protectedPath)) File.Delete(protectedPath);
    }

    private static async Task ProcessDocumentAsync(string inputPath, string outputPath, CancellationToken token)
    {
        // Load the document.
        Document doc = new Document(inputPath);

        // Simulate lengthy processing with periodic cancellation checks.
        for (int i = 0; i < 5; i++)
        {
            token.ThrowIfCancellationRequested();
            await Task.Delay(300, token); // Simulated work.
        }

        // Apply protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "pwd123");

        // Save with encryption (password protection) using OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "pwd123"
        };
        doc.Save(outputPath, saveOptions);
    }
}
