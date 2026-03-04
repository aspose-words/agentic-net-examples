using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image into the document.
        // The InsertImage method returns a Shape object that represents the image.
        Shape imageShape = builder.InsertImage("ImageDir/Logo.jpg");

        // Lock the shape's aspect ratio so that resizing preserves the original proportions.
        imageShape.AspectRatioLocked = true;

        // Save the document to the desired location.
        doc.Save("ArtifactsDir/AspectRatioLocked.docx");
    }
}
