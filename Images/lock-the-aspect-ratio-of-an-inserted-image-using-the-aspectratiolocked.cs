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

        // Insert an image into the document.
        // The InsertImage method returns a Shape object that represents the image.
        Shape imageShape = builder.InsertImage("ImageDir/Logo.jpg");

        // Lock the aspect ratio of the inserted image.
        // When true, the image will keep its original width‑to‑height proportion
        // when resized using diagonal handles in Microsoft Word.
        imageShape.AspectRatioLocked = true;

        // Save the document to the desired location.
        doc.Save("ArtifactsDir/AspectRatioLocked.docx");
    }
}
