using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Ensure the document has a paragraph to contain the shape.
        if (doc.FirstSection.Body.FirstParagraph == null)
        {
            doc.FirstSection.Body.AppendChild(new Paragraph(doc));
        }

        // Create a shape of type Image.
        Shape shape = new Shape(doc, ShapeType.Image);

        // Set the image file for the shape.
        shape.ImageData.SetImage("ImageDir/Windows MetaFile.wmf");

        // Define the shape size.
        shape.Width = 100;
        shape.Height = 100;

        // Append the shape to the first paragraph of the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

        // Save the document to a DOCX file.
        doc.Save("ArtifactsDir/Image.FromFile.docx");
    }
}
