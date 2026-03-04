using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file to be inserted.
        string imagePath = "ImageDir/Logo.jpg";

        // Open the image file as a read‑only stream.
        using (Stream imageStream = File.OpenRead(imagePath))
        {
            // Insert the image from the stream at the current cursor position.
            builder.InsertImage(imageStream);
        }

        // Save the resulting document.
        doc.Save("ArtifactsDir/InsertImageFromStream.docx");
    }
}
