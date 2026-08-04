using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Callback that checks the cancellation token and aborts the save operation.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly CancellationToken _cancellationToken;

        public SavingProgressCallback(CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
        }

        public void Notify(DocumentSavingArgs args)
        {
            if (_cancellationToken.IsCancellationRequested)
                throw new OperationCanceledException("Document saving was canceled via token.");
        }
    }

    public class Program
    {
        public static async Task Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "LargeDocument.docx");

            // Create a large document to make the save operation take noticeable time.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (int i = 0; i < 5000; i++)
            {
                builder.Writeln($"Paragraph {i + 1}");
            }

            // Set up a cancellation token that will be triggered after a short delay.
            using CancellationTokenSource cts = new CancellationTokenSource();
            _ = Task.Run(async () =>
            {
                await Task.Delay(200);
                cts.Cancel();
            });

            // Configure save options with a progress callback that respects the token.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback(cts.Token)
            };

            try
            {
                // Perform the save operation on a background thread so it can be cancelled.
                await Task.Run(() => doc.Save(outputPath, saveOptions), cts.Token);
                Console.WriteLine("Document saved successfully.");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Document saving was canceled.");
            }

            // Verify whether the file was created.
            if (File.Exists(outputPath))
                Console.WriteLine($"Output file exists at: {outputPath}");
            else
                Console.WriteLine("Output file was not created.");
        }
    }
}
