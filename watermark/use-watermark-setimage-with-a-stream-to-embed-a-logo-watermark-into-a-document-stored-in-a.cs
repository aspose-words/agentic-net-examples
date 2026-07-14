using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // 1. Prepare a simple PNG image (1x1 pixel, transparent) as a Base64 string.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Z6XcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using var imageStream = new MemoryStream(imageBytes);
        imageStream.Position = 0; // Ensure the stream is at the beginning.

        // 2. Create a blank Word document and add some content.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This document demonstrates adding an image watermark from a stream.");
        builder.Writeln("The watermark is a tiny PNG image embedded via a MemoryStream.");

        // 3. Simulate storing the document in Azure Blob storage by using a MemoryStream.
        using var blobStream = new MemoryStream();
        doc.Save(blobStream, SaveFormat.Docx);
        blobStream.Position = 0; // Reset for reading.

        // 4. Load the document from the simulated blob.
        var loadedDoc = new Document(blobStream);

        // 5. Apply the image watermark using the stream.
        var watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,
            IsWashout = false
        };
        imageStream.Position = 0; // Reset before reuse.
        loadedDoc.Watermark.SetImage(imageStream, watermarkOptions);

        // 6. Save the watermarked document to the local file system.
        const string outputPath = "WatermarkedDocument.docx";
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        Console.WriteLine($"Watermarked document saved to '{Path.GetFullPath(outputPath)}'.");
    }
}
