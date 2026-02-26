using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithImage
{
    static void Main()
    {
        // Path to the image that will be placed inside the shape.
        string imagePath = @"C:\Images\Logo.jpg";

        // Path where the resulting DOCX file will be saved.
        string outputPath = @"C:\Output\ImageShape.docx";

        // Create a new, empty Word document.
        Document doc = new Document();

        // Create a shape of type Image. The constructor automatically sets the markup language to VML.
        Shape shape = new Shape(doc, ShapeType.Image);

        // Load the image from the file system into the shape.
        shape.ImageData.SetImage(imagePath);

        // Set the desired size of the shape (in points). 1 point = 1/72 inch.
        shape.Width = 100;   // 100 points ≈ 1.39 inches
        shape.Height = 100;  // 100 points ≈ 1.39 inches

        // Append the shape to the first paragraph of the document body.
        // If the document has no paragraphs, FirstParagraph will be created automatically.
        doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

        // Save the document in DOCX format.
        doc.Save(outputPath);
    }
}
