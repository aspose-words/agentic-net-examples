using System;
using System.IO;
using System.Linq;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsImageExtraction
{
    class Program
    {
        static void Main()
        {
            // Path to the source document.
            const string docPath = @"C:\Docs\SourceDocument.docx";

            // Folder where extracted images will be saved.
            const string outputFolder = @"C:\Docs\ExtractedImages";

            // Ensure the output folder exists.
            Directory.CreateDirectory(outputFolder);

            // Create a cancellation token source that can be used to cancel the operation.
            using var cts = new CancellationTokenSource();

            // Example: cancel after 5 seconds (adjust as needed).
            // In a real scenario, cancellation could be triggered by UI or another thread.
            Timer timer = new Timer(_ => cts.Cancel(), null, TimeSpan.FromSeconds(5), Timeout.InfiniteTimeSpan);

            try
            {
                // Load the document using Aspose.Words (lifecycle rule: use provided load logic).
                Document doc = new Document(docPath);

                // Extract images with cancellation support.
                ExtractImages(doc, outputFolder, cts.Token);
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Image extraction was cancelled.");
            }
            finally
            {
                timer.Dispose();
            }
        }

        /// <summary>
        /// Extracts all images from the provided document and saves them to the specified folder.
        /// The method checks the cancellation token before processing each image and throws
        /// if cancellation has been requested.
        /// </summary>
        /// <param name="doc">The Aspose.Words document.</param>
        /// <param name="outputFolder">Folder to save extracted images.</param>
        /// <param name="token">Cancellation token to support early termination.</param>
        private static void ExtractImages(Document doc, string outputFolder, CancellationToken token)
        {
            // Get all shape nodes that may contain images.
            var shapes = doc.GetChildNodes(NodeType.Shape, true)
                            .Cast<Shape>()
                            .Where(s => s.HasImage)
                            .ToList();

            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                // Throw if cancellation was requested before processing the next image.
                token.ThrowIfCancellationRequested();

                // Determine a file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build a unique file name for the image.
                string imageFileName = Path.Combine(outputFolder,
                    $"Image_{imageIndex}{extension}");

                // Save the image to the file system.
                shape.ImageData.Save(imageFileName);

                Console.WriteLine($"Saved: {imageFileName}");

                imageIndex++;
            }
        }
    }
}
