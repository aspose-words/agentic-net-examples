using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image shape into the document.
        // Replace the path with the actual location of your image file.
        Shape shape = builder.InsertImage("ImageDir/Logo.jpg");

        // Lock the shape's aspect ratio so that resizing preserves its proportions.
        shape.AspectRatioLocked = true;

        // Save the resulting document.
        doc.Save("ArtifactsDir/Shape.AspectRatioLocked.docx");
    }
}
