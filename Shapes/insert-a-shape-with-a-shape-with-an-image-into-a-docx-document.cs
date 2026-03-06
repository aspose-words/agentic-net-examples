using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithImage
{
    static void Main()
    {
        // Path to the image file that will be placed inside the shape.
        string imagePath = @"C:\Images\SampleImage.jpg";

        // Create a new, empty Word document.
        Document doc = new Document();

        // Create a shape of type Image and associate it with the document.
        Shape shape = new Shape(doc, ShapeType.Image);

        // Load the image from the file system into the shape.
        shape.ImageData.SetImage(imagePath);

        // Set the desired size of the shape (in points).
        shape.Width = 100;
        shape.Height = 100;

        // Append the shape to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

        // Save the document to a DOCX file.
        doc.Save(@"C:\Output\ImageShape.docx");
    }
}
