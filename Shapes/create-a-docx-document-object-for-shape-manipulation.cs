using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a blank Word document (lifecycle rule)
        Document doc = new Document();

        // The default document already contains a Section, Body, and Paragraph.
        // Create a rectangle shape using the Shape constructor (rule)
        Shape shape = new Shape(doc, ShapeType.Rectangle);
        shape.Width = 100;               // Width in points
        shape.Height = 50;               // Height in points
        shape.Stroke.Color = Color.Blue; // Outline color
        shape.Fill.ForeColor = Color.LightGray; // Fill color

        // Append the shape to the first paragraph of the document
        doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

        // Save the document to a DOCX file (lifecycle rule)
        doc.Save("ShapeManipulation.docx");
    }
}
