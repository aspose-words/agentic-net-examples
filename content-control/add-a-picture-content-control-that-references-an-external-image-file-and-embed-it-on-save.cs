using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a picture content control (SDT) at the inline level.
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Inline)
        {
            Title = "SamplePicture",
            Tag = "sample-picture"
        };

        // Insert the content control into the document.
        builder.InsertNode(pictureSdt);

        // Prepare a simple 1x1 pixel PNG image (embedded as a base64 string).
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Z9VYAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create a shape of type Image and embed the PNG data using a MemoryStream.
        Shape pictureShape = new Shape(doc, ShapeType.Image);
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            pictureShape.ImageData.SetImage(ms);
        }
        pictureShape.Width = 50;   // Desired width (points).
        pictureShape.Height = 50;  // Desired height (points).

        // Add the picture shape as a child of the picture content control.
        pictureSdt.AppendChild(pictureShape);

        // Save the document; the image is embedded within the DOCX.
        doc.Save("PictureContentControl.docx");
    }
}
