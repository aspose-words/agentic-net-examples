using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedDocument.docx");

        // Use a CancellationTokenSource within a using block to ensure it is disposed.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Simulate a cancellation token check (not required by Aspose.Words but shown for completeness).
            CancellationToken token = cts.Token;

            // Create a new blank document.
            Document doc = new Document();

            // Build simple content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This document is protected with a password.");

            // Apply read‑only protection with a password.
            doc.Protect(ProtectionType.ReadOnly, "SecretPassword");

            // Save the protected document.
            doc.Save(outputPath, SaveFormat.Docx);

            // Optional: check for cancellation request.
            if (token.IsCancellationRequested)
            {
                // If cancellation was requested, exit early.
                return;
            }
        } // The CancellationTokenSource is disposed here, freeing system resources.

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The expected output file was not created: {outputPath}");
        }

        // Load the protected document to verify the protection state.
        Document loadedDoc = new Document(outputPath);
        if (loadedDoc.ProtectionType != ProtectionType.ReadOnly)
        {
            throw new InvalidOperationException("The document protection was not applied as expected.");
        }

        // Unprotect the document (no password needed for Aspose.Words Unprotect method).
        loadedDoc.Unprotect();

        // Save the unprotected version.
        string unprotectedPath = Path.Combine(Directory.GetCurrentDirectory(), "UnprotectedDocument.docx");
        loadedDoc.Save(unprotectedPath, SaveFormat.Docx);

        // Final validation.
        if (!File.Exists(unprotectedPath))
        {
            throw new InvalidOperationException($"Failed to save the unprotected document: {unprotectedPath}");
        }
    }
}
