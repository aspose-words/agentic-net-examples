using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Extension method that adds cancellation support to Document.Save.
    public static class DocumentExtensions
    {
        public static void Save(this Document doc, string fileName, CancellationToken cancellationToken)
        {
            // Throw immediately if cancellation was already requested.
            if (cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException(cancellationToken);

            // Use OoxmlSaveOptions so we can attach a progress callback.
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingCancellationCallback(cancellationToken)
            };

            // Perform the synchronous save; the callback will abort if the token is cancelled.
            doc.Save(fileName, saveOptions);
        }
    }

    // Callback that checks the cancellation token and aborts the save operation.
    internal class SavingCancellationCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _token;

        public SavingCancellationCallback(CancellationToken token) => _token = token;

        public void Notify(DocumentSavingArgs args)
        {
            if (_token.IsCancellationRequested)
                throw new OperationCanceledException(_token);
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a simple document.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words with cancellation support!");

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

            // Use a non‑cancelled token for this demo.
            var cancellationToken = CancellationToken.None;

            // Save the document using the extension method.
            doc.Save(outputPath, cancellationToken);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved as expected.");

            // Optionally, clean up the file after verification.
            File.Delete(outputPath);
        }
    }
}
