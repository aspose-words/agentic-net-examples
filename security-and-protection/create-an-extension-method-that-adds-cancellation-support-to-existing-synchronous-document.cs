using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeCancellationDemo
{
    // Extension method that adds cancellation support to Document.Save
    public static class DocumentExtensions
    {
        public static void Save(this Document doc, string fileName, CancellationToken cancellationToken)
        {
            // Determine the save format from the file extension
            SaveFormat format = GetSaveFormatFromExtension(Path.GetExtension(fileName));

            // Create concrete SaveOptions for the detected format and attach a progress callback that checks the token
            SaveOptions saveOptions = GetSaveOptions(format, cancellationToken);

            // Perform the save using the options
            doc.Save(fileName, saveOptions);
        }

        private static SaveFormat GetSaveFormatFromExtension(string extension)
        {
            return extension.ToLower() switch
            {
                ".docx" => SaveFormat.Docx,
                ".doc"  => SaveFormat.Doc,
                ".pdf"  => SaveFormat.Pdf,
                ".html" => SaveFormat.Html,
                ".txt"  => SaveFormat.Text,
                ".odt"  => SaveFormat.Odt,
                _       => SaveFormat.Docx,
            };
        }

        // Returns a concrete SaveOptions instance appropriate for the given format
        private static SaveOptions GetSaveOptions(SaveFormat format, CancellationToken token)
        {
            // Helper to assign the cancellation callback
            SaveOptions AttachCallback(SaveOptions options)
            {
                options.ProgressCallback = new CancellationSavingCallback(token);
                return options;
            }

            return format switch
            {
                SaveFormat.Docx or SaveFormat.Docm or SaveFormat.Dotx or SaveFormat.Dotm
                    => AttachCallback(new OoxmlSaveOptions(format)),

                SaveFormat.Doc or SaveFormat.Dot
                    => AttachCallback(new DocSaveOptions(format)),

                SaveFormat.Pdf
                    => AttachCallback(new PdfSaveOptions()),

                SaveFormat.Html
                    => AttachCallback(new HtmlSaveOptions(format)),

                SaveFormat.Text
                    => AttachCallback(new TxtSaveOptions()),

                SaveFormat.Odt
                    => AttachCallback(new OdtSaveOptions(format)),

                _ => AttachCallback(new OoxmlSaveOptions(SaveFormat.Docx)),
            };
        }

        // Callback that aborts saving when the token is cancelled
        private class CancellationSavingCallback : IDocumentSavingCallback
        {
            private readonly CancellationToken _token;
            public CancellationSavingCallback(CancellationToken token) => _token = token;

            public void Notify(DocumentSavingArgs args)
            {
                if (_token.IsCancellationRequested)
                    throw new OperationCanceledException("Document saving was canceled.");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory
            string outDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
            Directory.CreateDirectory(outDir);
            string filePath = Path.Combine(outDir, "Sample.docx");

            // Create a simple document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello Aspose.Words with cancellation support!");

            // Save without cancellation
            var cts = new CancellationTokenSource();
            doc.Save(filePath, cts.Token);
            if (!File.Exists(filePath))
                throw new Exception("File was not saved as expected.");

            // Attempt to save with a pre‑cancelled token
            string cancelledPath = Path.Combine(outDir, "Cancelled.docx");
            var cancelledCts = new CancellationTokenSource();
            cancelledCts.Cancel(); // cancel before invoking save

            try
            {
                doc.Save(cancelledPath, cancelledCts.Token);
                throw new Exception("Save should have been canceled but completed.");
            }
            catch (OperationCanceledException)
            {
                // Expected outcome – saving was aborted.
            }
        }
    }
}
