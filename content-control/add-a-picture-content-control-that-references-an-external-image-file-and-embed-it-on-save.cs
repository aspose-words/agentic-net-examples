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

        // Prepare a tiny PNG image (1x1 pixel) as a byte array.
        // This ensures the example is self‑contained and does not rely on external files.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // Insert a picture content control at the current cursor position.
        StructuredDocumentTag pictureSdt = builder.InsertStructuredDocumentTag(SdtType.Picture);
        pictureSdt.Title = "SamplePicture";
        pictureSdt.Tag = "sample-picture";

        // Create a Shape of type Image, load the PNG, and add it as a child of the picture content control.
        Shape pictureShape = new Shape(doc, ShapeType.Image);
        pictureShape.ImageData.SetImage(imagePath);
        pictureShape.WrapType = WrapType.Inline;

        // A picture content control must contain a paragraph before a shape.
        Paragraph innerParagraph = new Paragraph(doc);
        pictureSdt.AppendChild(innerParagraph);
        innerParagraph.AppendChild(pictureShape);

        // Save the document. The image is embedded because SetImage copies the image data into the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PictureContentControl.docx");
        doc.Save(outputPath);

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
        {
            File.Delete(imagePath);
        }

        Console.WriteLine("Document saved to: " + outputPath);
    }
}
