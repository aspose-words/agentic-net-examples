using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ImageInserter
{
    // Inserts each image from the byte array collection into the document sequentially.
    // After each image a page break is added to start the next image on a new page.
    public static void InsertImages(IEnumerable<byte[]> imageBytesCollection, string outputFilePath)
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Iterate over the collection of image byte arrays.
        foreach (byte[] imageBytes in imageBytesCollection)
        {
            // Insert the image inline with its original dimensions.
            // This uses the InsertImage overload that accepts a byte[].
            builder.InsertImage(imageBytes);

            // Insert a page break so the next image starts on a new page.
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the resulting document to the specified path.
        doc.Save(outputFilePath);
    }

    // Example usage.
    public static void Main()
    {
        // Small 1x1 pixel PNG image (red) encoded in base64.
        const string base64Png = 
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";

        // Prepare a list of image byte arrays.
        List<byte[]> images = new List<byte[]>
        {
            Convert.FromBase64String(base64Png),
            Convert.FromBase64String(base64Png),
            Convert.FromBase64String(base64Png)
        };

        // Define the output document path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MultipleImages.docx");

        // Insert the images into the document.
        InsertImages(images, outputPath);

        Console.WriteLine("Document created at: " + outputPath);
    }
}
