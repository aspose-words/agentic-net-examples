using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to host the shape.
        // If the document is empty, Aspose.Words creates a first paragraph automatically.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;

        // Create a shape of type Image. This shape will hold the picture.
        Shape pictureShape = new Shape(doc, ShapeType.Image);

        // Set the image that the shape will display.
        // Replace the path with the actual location of your image file.
        string imagePath = Path.Combine(Environment.CurrentDirectory, "SampleImage.jpg");
        pictureShape.ImageData.SetImage(imagePath);

        // Optionally set the size of the shape (in points).
        pictureShape.Width = 200;   // 200 points wide
        pictureShape.Height = 150;  // 150 points high

        // Append the shape to the paragraph so it becomes part of the document.
        para.AppendChild(pictureShape);

        // Save the document in DOCX format.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ShapeWithPicture.docx");
        doc.Save(outputPath);
    }
}
