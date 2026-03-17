using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class ChangePictureToAutoShape
{
    static void Main()
    {
        // Load the existing document that contains a picture shape.
        Document doc = new Document(@"C:\Docs\Input.docx");

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

        // Save the modified document.
        doc.Save(@"C:\Docs\Output.docx");
    }
}
