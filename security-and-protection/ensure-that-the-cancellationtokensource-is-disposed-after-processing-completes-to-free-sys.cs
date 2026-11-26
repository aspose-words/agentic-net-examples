using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample documents.
        const string sourcePath = "Sample.docx";
        const string protectedPath = "Protected.docx";

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World");
        doc.Save(sourcePath);

        // Apply read‑only protection with a password.
        doc.Protect(ProtectionType.ReadOnly, "myPassword");
        doc.Save(protectedPath);

        // Verify that the protected file was created.
        if (!File.Exists(protectedPath))
            throw new FileNotFoundException("Protected document was not saved.", protectedPath);

        // Use a CancellationTokenSource to load the protected document.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            CancellationToken token = cts.Token;

            // Load the protected document in a separate task, respecting the token.
            Task loadTask = Task.Run(() =>
            {
                // LoadOptions with the password required to open the protected file.
                LoadOptions loadOptions = new LoadOptions { Password = "myPassword" };
                Document loadedDoc = new Document(protectedPath, loadOptions);

                // Simple validation: the document should contain the original text.
                if (!loadedDoc.GetText().Contains("Hello World"))
                    throw new InvalidOperationException("Loaded document content is unexpected.");
            }, token);

            // Wait for the task to complete.
            loadTask.Wait();
        }

        // At this point the CancellationTokenSource has been disposed,
        // freeing any associated system resources.
        Console.WriteLine("Document processed and CancellationTokenSource disposed successfully.");
    }
}
