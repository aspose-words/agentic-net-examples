using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Open the image file as a read‑only stream.
        using (Stream imageStream = File.OpenRead("ImageDir/Logo.jpg"))
        {
            // Insert the image from the stream into the document.
            // The image is inserted inline at its original size (100% scale).
            builder.InsertImage(imageStream);
        }

        // Save the resulting document to disk.
        doc.Save("ArtifactsDir/DocumentWithImageFromStream.docx");
    }
}
