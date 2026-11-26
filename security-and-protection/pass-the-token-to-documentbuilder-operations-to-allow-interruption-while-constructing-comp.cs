using System;
using System.IO;
using System.Threading;
using System.Timers;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Set up a cancellation token that will be triggered after a short delay.
        var cts = new CancellationTokenSource();
        var timer = new System.Timers.Timer(100);
        timer.Elapsed += (s, e) =>
        {
            cts.Cancel();
            timer.Stop();
        };
        timer.Start();

        // Create a new document and a builder.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Build a complex document, checking the token to allow interruption.
        for (int i = 0; i < 20; i++)
        {
            if (cts.Token.IsCancellationRequested)
            {
                Console.WriteLine($"Construction interrupted at paragraph {i + 1}.");
                break;
            }

            builder.Writeln($"Paragraph {i + 1}: Sample text for demonstration purposes.");
        }

        // Apply read‑only protection with a password.
        const string password = "Secret123";
        doc.Protect(ProtectionType.ReadOnly, password);

        // Save the protected document.
        const string outputPath = "ProtectedDocument.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the protected document.");

        // Reopen the document to confirm protection.
        var loadedDoc = new Document(outputPath);
        if (loadedDoc.ProtectionType != ProtectionType.ReadOnly)
            throw new InvalidOperationException("Document protection verification failed.");

        Console.WriteLine("Document created, protected, and verified successfully.");
    }
}
