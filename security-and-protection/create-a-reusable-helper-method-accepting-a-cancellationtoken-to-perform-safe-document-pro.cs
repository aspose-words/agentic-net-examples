using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    // Reusable helper that performs document creation, protection, saving, loading and unprotection.
    // The method respects the provided CancellationToken and throws if cancellation is requested.
    public static async Task SafeProcessDocumentAsync(CancellationToken cancellationToken)
    {
        // Define file paths in the temporary folder.
        string tempFolder = Path.GetTempPath();
        string protectedPath = Path.Combine(tempFolder, "ProtectedDocument.docx");
        string unprotectedPath = Path.Combine(tempFolder, "UnprotectedDocument.docx");
        const string password = "Secret";

        // Step 1: Create a new blank document and add some text.
        cancellationToken.ThrowIfCancellationRequested();
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This is a protected document.");

        // Step 2: Apply read‑only protection with a password.
        cancellationToken.ThrowIfCancellationRequested();
        doc.Protect(ProtectionType.ReadOnly, password);

        // Step 3: Save the protected document to disk.
        cancellationToken.ThrowIfCancellationRequested();
        doc.Save(protectedPath);

        // Validate that the file was created.
        if (!File.Exists(protectedPath))
            throw new InvalidOperationException("Protected document was not saved correctly.");

        // Step 4: Load the protected document using the correct password.
        cancellationToken.ThrowIfCancellationRequested();
        LoadOptions loadOptions = new LoadOptions(password);
        Document loadedDoc = new Document(protectedPath, loadOptions);

        // Verify that the protection type is as expected.
        if (loadedDoc.ProtectionType != ProtectionType.ReadOnly)
            throw new InvalidOperationException("Loaded document does not have the expected protection.");

        // Step 5: Remove protection.
        cancellationToken.ThrowIfCancellationRequested();
        bool unprotected = loadedDoc.Unprotect(password);
        if (!unprotected || loadedDoc.ProtectionType != ProtectionType.NoProtection)
            throw new InvalidOperationException("Failed to unprotect the document.");

        // Step 6: Save the unprotected version.
        cancellationToken.ThrowIfCancellationRequested();
        loadedDoc.Save(unprotectedPath);

        // Final validation.
        if (!File.Exists(unprotectedPath))
            throw new InvalidOperationException("Unprotected document was not saved correctly.");
    }

    public static void Main()
    {
        // Create a CancellationTokenSource with a generous timeout (e.g., 30 seconds).
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try
        {
            // Run the helper synchronously for this console app.
            SafeProcessDocumentAsync(cts.Token).GetAwaiter().GetResult();
            Console.WriteLine("Document processing completed successfully.");
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
