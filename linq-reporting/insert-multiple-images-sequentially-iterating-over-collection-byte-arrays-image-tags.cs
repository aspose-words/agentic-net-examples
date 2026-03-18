using System;
using System.Collections.Generic;
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
        // Prepare a list of image byte arrays (replace with actual image data).
        List<byte[]> images = new List<byte[]>
        {
            System.IO.File.ReadAllBytes(@"C:\Images\Image1.jpg"),
            System.IO.File.ReadAllBytes(@"C:\Images\Image2.png"),
            System.IO.File.ReadAllBytes(@"C:\Images\Image3.bmp")
        };

        // Define the output document path.
        string outputPath = @"C:\Output\MultipleImages.docx";

        // Insert the images into the document.
        InsertImages(images, outputPath);

        Console.WriteLine("Document created at: " + outputPath);
    }
}
