using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;   // Needed for LoadOptions
using Aspose.Words.Saving;

namespace AsposeWordsHelperExample
{
    public class Program
    {
        // Reusable helper that safely processes a document respecting a CancellationToken.
        // It loads a password‑protected document, removes protection, appends text, and saves it.
        public static void SafeProcessDocument(string sourcePath, string destinationPath, string password, CancellationToken cancellationToken)
        {
            // Throw if cancellation was requested before starting.
            cancellationToken.ThrowIfCancellationRequested();

            // Load the protected document using the supplied password.
            LoadOptions loadOptions = new LoadOptions(password);
            Document doc = new Document(sourcePath, loadOptions);

            // Check for cancellation again after loading.
            cancellationToken.ThrowIfCancellationRequested();

            // Unprotect the document using the correct password.
            bool unprotected = doc.Unprotect(password);
            if (!unprotected)
                throw new InvalidOperationException("Failed to unprotect the document with the provided password.");

            // Check for cancellation before modifying the document.
            cancellationToken.ThrowIfCancellationRequested();

            // Append a new paragraph.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Appended text added safely.");

            // Check for cancellation before saving.
            cancellationToken.ThrowIfCancellationRequested();

            // Save the processed document.
            doc.Save(destinationPath);

            // Validate that the output file was created.
            if (!File.Exists(destinationPath))
                throw new FileNotFoundException("The processed document was not saved.", destinationPath);
        }

        public static void Main()
        {
            // Define file paths in the local execution directory.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            string sourcePath = Path.Combine(artifactsDir, "ProtectedDocument.docx");
            string destinationPath = Path.Combine(artifactsDir, "ProcessedDocument.docx");
            string password = "SecretPwd";

            // Create a sample source document and protect it with a password.
            Document sourceDoc = new Document();
            DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
            sourceBuilder.Writeln("Original content of the protected document.");
            sourceDoc.Protect(ProtectionType.ReadOnly, password);
            sourceDoc.Save(sourcePath);

            // Prepare a CancellationToken that is not cancelled.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                // Perform safe processing.
                SafeProcessDocument(sourcePath, destinationPath, password, cts.Token);
            }

            // Simple verification that the processed file exists.
            if (File.Exists(destinationPath))
            {
                Console.WriteLine("Document processed and saved successfully at:");
                Console.WriteLine(destinationPath);
            }
            else
            {
                Console.WriteLine("Processing failed; output file not found.");
            }
        }
    }
}
