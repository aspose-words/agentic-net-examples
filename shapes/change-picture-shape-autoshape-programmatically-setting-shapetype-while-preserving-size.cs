using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ChangePictureToAutoShape
{
    static void Main()
    {
        // Create a new document and insert a tiny picture shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 1x1 pixel PNG (transparent) as base64.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6eUAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        // Find the first shape that is an image (picture).
        Shape pictureShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (pictureShape == null || pictureShape.ShapeType != ShapeType.Image)
        {
            Console.WriteLine("No picture shape found in the document.");
            return;
        }

        // Preserve the original size and layout properties.
        double width = pictureShape.Width;
        double height = pictureShape.Height;
        double left = pictureShape.Left;
        double top = pictureShape.Top;
        RelativeHorizontalPosition horzPos = pictureShape.RelativeHorizontalPosition;
        RelativeVerticalPosition vertPos = pictureShape.RelativeVerticalPosition;
        WrapType wrap = pictureShape.WrapType;
        WrapSide wrapSide = pictureShape.WrapSide;
        HorizontalAlignment hAlign = pictureShape.HorizontalAlignment;
        VerticalAlignment vAlign = pictureShape.VerticalAlignment;

        // Create a new AutoShape (e.g., a rectangle) with the same dimensions.
        Shape autoShape = new Shape(pictureShape.Document, ShapeType.Rectangle);
        autoShape.Width = width;
        autoShape.Height = height;
        autoShape.Left = left;
        autoShape.Top = top;
        autoShape.RelativeHorizontalPosition = horzPos;
        autoShape.RelativeVerticalPosition = vertPos;
        autoShape.WrapType = wrap;
        autoShape.WrapSide = wrapSide;
        autoShape.HorizontalAlignment = hAlign;
        autoShape.VerticalAlignment = vAlign;

        // Insert the new shape after the original picture and then remove the picture.
        pictureShape.ParentNode.InsertAfter(autoShape, pictureShape);
        pictureShape.Remove();

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
