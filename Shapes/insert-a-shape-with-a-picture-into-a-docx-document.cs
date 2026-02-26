using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithPicture
{
    static void Main()
    {
        // Path to the image file that will be placed inside the shape.
        string imagePath = @"C:\Images\picture.jpg";

        // Path where the resulting DOCX document will be saved.
        string outputPath = @"C:\Output\ShapeWithImage.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline shape of type Image with desired size (in points).
        // Width and height are set to 150 points each (approximately 2 inches).
        Shape shape = builder.InsertShape(ShapeType.Image, 150, 150);

        // Load the image from the file system into the shape.
        shape.ImageData.SetImage(imagePath);

        // Optionally, set the shape's wrap type or positioning here.
        // shape.WrapType = WrapType.Inline; // default for InsertShape

        // Save the document in DOCX format.
        doc.Save(outputPath);
    }
}
