using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithPicture
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new shape of type Image.
        Shape pictureShape = new Shape(doc, ShapeType.Image);

        // Load an image file into the shape.
        // Replace "MyImage.jpg" with the actual path to your picture.
        pictureShape.ImageData.SetImage("MyImage.jpg");

        // Optionally set the size of the shape (in points).
        pictureShape.Width = 200;   // 200 points ≈ 2.78 inches
        pictureShape.Height = 150;  // 150 points ≈ 2.08 inches

        // Insert the shape into the document at the current builder position.
        builder.InsertNode(pictureShape);

        // Save the document to a DOCX file.
        doc.Save("ShapeWithPicture.docx");
    }
}
