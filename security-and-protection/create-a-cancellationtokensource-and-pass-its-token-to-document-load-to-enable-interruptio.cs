using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a temporary DOCX file to work with.
        string filePath = Path.Combine(Path.GetTempPath(), "Sample.docx");
        Document doc = new Document();          // empty document
        doc.Save(filePath);                     // save it so we have something to load later

        // Create a CancellationTokenSource and cancel it immediately.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            cts.Cancel(); // request cancellation before the load starts

            // Prepare load options (no special settings needed for this example).
            LoadOptions loadOptions = new LoadOptions();

            try
            {
                // Manually respect the cancellation token before invoking the load operation.
                // Aspose.Words does not expose a Load overload that accepts a token in this version,
                // so we check the token ourselves and throw if cancellation was requested.
                if (cts.Token.IsCancellationRequested)
                    throw new OperationCanceledException(cts.Token);

                // Load the document normally.
                Document loadedDoc = new Document(filePath, loadOptions);
                Console.WriteLine("Document loaded successfully (unexpected).");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Document loading was cancelled as expected.");
            }
        }

        // Clean up the temporary file.
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
}
