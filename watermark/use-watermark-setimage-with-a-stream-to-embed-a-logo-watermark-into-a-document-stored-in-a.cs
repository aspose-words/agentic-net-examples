using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // 1. Prepare a simple 1x1 pixel PNG image (transparent) as a byte array.
        // This avoids any external image files or System.Drawing usage.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Z6YcAAAAASUVORK5CYII=");

        // 2. Create a blank document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an image watermark applied from a stream.");

        // 3. Simulate uploading the document to Azure Blob Storage by saving it to a memory stream.
        using (MemoryStream blobStream = new MemoryStream())
        {
            doc.Save(blobStream, SaveFormat.Docx);
            // Reset the stream position to the beginning to simulate downloading.
            blobStream.Position = 0;

            // 4. Simulate downloading the document from Azure Blob Storage.
            Document loadedDoc = new Document(blobStream);

            // 5. Prepare the image stream for the watermark.
            using (MemoryStream imageStream = new MemoryStream(pngBytes))
            {
                // Ensure the stream is at the beginning.
                imageStream.Position = 0;

                // 6. Configure watermark options (optional).
                ImageWatermarkOptions options = new ImageWatermarkOptions
                {
                    Scale = 5,          // Example scale factor.
                    IsWashout = false   // Show the image without washout effect.
                };

                // 7. Apply the image watermark using the stream overload.
                loadedDoc.Watermark.SetImage(imageStream, options);
            }

            // 8. Save the resulting document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputWithWatermark.docx");
            loadedDoc.Save(outputPath);
        }

        // Indicate completion (no interactive prompts).
        Console.WriteLine("Document with image watermark has been created.");
    }
}
