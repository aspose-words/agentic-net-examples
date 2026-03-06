using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithImage
{
    static void Main()
    {
        // Path to the image file that will be placed inside the shape.
        string imageFile = @"C:\Images\Logo.jpg";

        // Path where the resulting DOCX document will be saved.
        string outputFile = @"C:\Output\ImageShape.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Create a shape of type Image and associate it with the document.
        Shape shape = new Shape(doc, ShapeType.Image);

        // Load the image from the file system into the shape.
        shape.ImageData.SetImage(imageFile);

        // Set the desired size of the shape (in points).
        shape.Width = 100;
        shape.Height = 100;

        // Append the shape to the first paragraph of the first section.
        doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

        // Save the document in DOCX format.
        doc.Save(outputFile);
    }
}
