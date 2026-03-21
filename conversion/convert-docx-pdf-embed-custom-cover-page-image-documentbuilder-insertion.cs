using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the very beginning of the document.
            builder.MoveToDocumentStart();

            // Simple PNG image (1x1 pixel) as a placeholder cover page.
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
            using (MemoryStream imageStream = new MemoryStream(pngBytes))
            {
                // Insert the cover page image. The image is inserted inline at 100% scale.
                builder.InsertImage(imageStream);
            }

            // Add a sample paragraph after the cover image.
            builder.Writeln("\nThis is the main document content.");

            // Determine output PDF path in the current directory.
            string outputPdf = Path.Combine(Environment.CurrentDirectory, "ResultDocument.pdf");

            // Save the modified document as PDF.
            doc.Save(outputPdf, SaveFormat.Pdf);

            Console.WriteLine($"PDF saved to: {outputPdf}");
        }
    }
}
