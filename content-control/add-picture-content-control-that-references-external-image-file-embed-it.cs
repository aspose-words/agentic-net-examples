using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Path to the external image file.
        string imagePath = Path.Combine(Environment.CurrentDirectory, "Sample.jpg");
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image not found: {imagePath}");
            return;
        }

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a picture content control.
        StructuredDocumentTag pictureControl = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Block);
        pictureControl.Title = "ExternalPictureControl";

        // Insert the content control at the current builder position.
        builder.InsertNode(pictureControl);

        // Move the builder inside the content control and insert the image.
        builder.MoveTo(pictureControl);
        builder.InsertImage(imagePath);

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DocumentWithEmbeddedPicture.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
