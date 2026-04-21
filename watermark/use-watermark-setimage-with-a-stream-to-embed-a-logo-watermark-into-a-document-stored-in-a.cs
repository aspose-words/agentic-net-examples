using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Base64 string for a minimal 1x1 PNG image.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create a simple Word document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("The logo watermark will appear behind this text.");

        // Simulate storing the document in Azure Blob storage using a memory stream.
        using (MemoryStream blobStream = new MemoryStream())
        {
            originalDoc.Save(blobStream, SaveFormat.Docx);
            blobStream.Position = 0; // Reset for reading.

            // Load the document from the simulated blob.
            Document docFromBlob = new Document(blobStream);

            // Apply an image watermark using a fresh stream.
            using (MemoryStream watermarkImageStream = new MemoryStream(imageBytes))
            {
                ImageWatermarkOptions options = new ImageWatermarkOptions
                {
                    IsWashout = false, // No washout effect.
                    Scale = 5           // Increase watermark size.
                };

                docFromBlob.Watermark.SetImage(watermarkImageStream, options);
            }

            // Save the watermarked document to the local file system.
            const string outputPath = "WatermarkedDocument.docx";
            docFromBlob.Save(outputPath, SaveFormat.Docx);

            // Verify that the file was created.
            Console.WriteLine(File.Exists(outputPath)
                ? $"Watermarked document saved successfully to '{Path.GetFullPath(outputPath)}'."
                : "Failed to save the watermarked document.");
        }
    }
}
