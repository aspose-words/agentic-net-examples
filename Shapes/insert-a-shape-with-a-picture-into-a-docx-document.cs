using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithPicture
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document (optional for this task).
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file that will be placed inside the shape.
        string imagePath = @"C:\Images\SampleImage.jpg";

        // Create a shape of type Image and associate it with the document.
        Shape pictureShape = new Shape(doc, ShapeType.Image);

        // Load the image from the file system into the shape.
        pictureShape.ImageData.SetImage(imagePath);

        // Set desired dimensions for the shape (in points).
        pictureShape.Width = 150;   // 150 points width
        pictureShape.Height = 100;  // 100 points height

        // Append the shape to the first paragraph of the document body.
        doc.FirstSection.Body.FirstParagraph.AppendChild(pictureShape);

        // Save the document as a DOCX file.
        doc.Save(@"C:\Output\ShapeWithPicture.docx");
    }
}
