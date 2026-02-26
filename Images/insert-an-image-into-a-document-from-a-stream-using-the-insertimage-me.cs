using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertImageFromStream
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file that will be read into a stream.
        string imagePath = @"C:\Images\Logo.jpg";

        // Open the image file as a read‑only stream.
        using (Stream imageStream = File.OpenRead(imagePath))
        {
            // Insert the image from the stream into the document.
            // The image is inserted inline at its original size.
            Shape insertedShape = builder.InsertImage(imageStream);

            // Optional: modify the inserted shape (e.g., set wrap type).
            insertedShape.WrapType = WrapType.Inline;
        }

        // Save the document to disk.
        string outputPath = @"C:\Output\ImageFromStream.docx";
        doc.Save(outputPath);
    }
}
