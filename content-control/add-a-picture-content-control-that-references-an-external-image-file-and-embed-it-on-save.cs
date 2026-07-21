using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Prepare a simple 1x1 PNG image file in the working directory.
        const string imageFileName = "sample.png";
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcVYAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imageFileName, imageBytes);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a picture content control (inline level).
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Inline)
        {
            Title = "PictureControl",
            Tag = "pic-1"
        };

        // Insert the content control into the document.
        builder.InsertNode(pictureSdt);

        // Move the builder inside the newly inserted content control.
        builder.MoveTo(pictureSdt);

        // Insert the image. The image is embedded into the document.
        builder.InsertImage(imageFileName);

        // Save the resulting document.
        doc.Save("PictureContentControl.docx");
    }
}
