using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsSecurityExample
{
    public class Program
    {
        // Reusable helper that safely processes a document with cancellation support.
        // It loads a password‑protected document, removes protection, and saves the result.
        public static void ProcessDocument(string inputPath, string outputPath, string password, CancellationToken cancellationToken)
        {
            // Throw if cancellation was requested before starting.
            cancellationToken.ThrowIfCancellationRequested();

            // Load the protected document using the supplied password.
            var loadOptions = new LoadOptions
            {
                Password = password
            };
            Document doc = new Document(inputPath, loadOptions);

            // Check for cancellation again after loading.
            cancellationToken.ThrowIfCancellationRequested();

            // Remove protection if present.
            if (doc.ProtectionType != ProtectionType.NoProtection)
            {
                doc.Unprotect(password);
            }

            // Save the unprotected document.
            doc.Save(outputPath);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }

        public static void Main()
        {
            // Create a simple source document.
            string sourcePath = "source.docx";
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);
            builder.Writeln("This is a sample document for security demonstration.");
            sourceDoc.Save(sourcePath);

            // Apply read‑only protection with a password.
            string protectedPath = "protected.docx";
            string password = "Secret123";
            Document protectedDoc = new Document(sourcePath);
            protectedDoc.Protect(ProtectionType.ReadOnly, password);
            protectedDoc.Save(protectedPath);

            // Prepare cancellation token (not cancelled in this example).
            using var cts = new CancellationTokenSource();
            CancellationToken token = cts.Token;

            // Process the protected document: remove protection and save unprotected copy.
            string unprotectedPath = "unprotected.docx";
            ProcessDocument(protectedPath, unprotectedPath, password, token);

            // Simple verification that the unprotected document is indeed not protected.
            Document resultDoc = new Document(unprotectedPath);
            if (resultDoc.ProtectionType != ProtectionType.NoProtection)
                throw new InvalidOperationException("The document is still protected after processing.");

            // Clean up temporary files (optional).
            File.Delete(sourcePath);
            File.Delete(protectedPath);
            File.Delete(unprotectedPath);
        }
    }
}
