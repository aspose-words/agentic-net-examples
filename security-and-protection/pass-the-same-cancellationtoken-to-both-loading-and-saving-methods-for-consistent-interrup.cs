using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static async Task Main()
    {
        // Create temporary folder for files
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeDemo_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);

        string originalPath = Path.Combine(tempFolder, "Original.docx");
        string protectedPath = Path.Combine(tempFolder, "Protected.docx");
        string unprotectedPath = Path.Combine(tempFolder, "Unprotected.docx");

        // Build a simple document and save it
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        doc.Save(originalPath, new OoxmlSaveOptions());

        // Create a single CancellationTokenSource to be used for both load and save
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Load the original document using the shared token
        Document loadedDoc = await Task.Run(() =>
        {
            token.ThrowIfCancellationRequested();
            return new Document(originalPath, new LoadOptions());
        }, token);

        // Apply read‑only protection with a password
        loadedDoc.Protect(ProtectionType.ReadOnly, "SecretPassword");

        // Save the protected document using the same token
        await Task.Run(() =>
        {
            token.ThrowIfCancellationRequested();
            loadedDoc.Save(protectedPath, new OoxmlSaveOptions());
        }, token);

        // Verify the protected file was created
        if (!File.Exists(protectedPath))
            throw new InvalidOperationException("Protected document was not saved.");

        // Load the protected document with the correct password, again using the shared token
        var loadOptions = new LoadOptions { Password = "SecretPassword" };
        Document protectedDoc = await Task.Run(() =>
        {
            token.ThrowIfCancellationRequested();
            return new Document(protectedPath, loadOptions);
        }, token);

        // Remove protection
        protectedDoc.Unprotect("SecretPassword");

        // Save the unprotected version using the same token
        await Task.Run(() =>
        {
            token.ThrowIfCancellationRequested();
            protectedDoc.Save(unprotectedPath, new OoxmlSaveOptions());
        }, token);

        // Verify the unprotected file was created
        if (!File.Exists(unprotectedPath))
            throw new InvalidOperationException("Unprotected document was not saved.");

        // Optional cleanup
        // Directory.Delete(tempFolder, true);
    }
}
