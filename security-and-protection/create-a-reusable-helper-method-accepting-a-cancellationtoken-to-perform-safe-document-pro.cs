using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    // Helper method that performs document processing safely, respecting cancellation.
    public static async Task<string> ProcessDocumentAsync(CancellationToken cancellationToken)
    {
        // Run the processing on a background thread to allow cancellation checks.
        return await Task.Run(() =>
        {
            // Check for cancellation before starting.
            if (cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(cancellationToken);

            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a new blank document and add some text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello Aspose.Words! This document will be write‑protected.");

            // Check for cancellation.
            if (cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(cancellationToken);

            // 2. Apply write protection with a password.
            doc.WriteProtection.SetPassword("SecretPwd");
            doc.WriteProtection.ReadOnlyRecommended = true;

            // 3. Save the protected document.
            string protectedPath = Path.Combine(outputDir, "Protected.docx");
            doc.Save(protectedPath);

            // Validate that the file was created.
            if (!File.Exists(protectedPath))
                throw new InvalidOperationException("Protected document was not saved.");

            // Check for cancellation.
            if (cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(cancellationToken);

            // 4. Load the protected document (write protection does not encrypt the file,
            //    so no password is needed for loading).
            Document loadedDoc = new Document(protectedPath, new LoadOptions());

            // Verify that write protection is still active.
            if (!loadedDoc.WriteProtection.IsWriteProtected ||
                !loadedDoc.WriteProtection.ValidatePassword("SecretPwd"))
                throw new InvalidOperationException("Write protection validation failed.");

            // 5. Remove write protection.
            loadedDoc.WriteProtection.SetPassword(string.Empty);
            loadedDoc.WriteProtection.ReadOnlyRecommended = false;

            // 6. Save the unprotected version.
            string unprotectedPath = Path.Combine(outputDir, "Unprotected.docx");
            loadedDoc.Save(unprotectedPath);

            // Validate that the unprotected file exists.
            if (!File.Exists(unprotectedPath))
                throw new InvalidOperationException("Unprotected document was not saved.");

            // Return the path of the final document.
            return unprotectedPath;
        }, cancellationToken);
    }

    public static void Main()
    {
        // Use a CancellationToken that is not cancelled to let the process complete.
        CancellationTokenSource cts = new CancellationTokenSource();

        try
        {
            string resultPath = ProcessDocumentAsync(cts.Token).GetAwaiter().GetResult();
            Console.WriteLine($"Document processing completed. Output file: {resultPath}");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document processing was canceled.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
