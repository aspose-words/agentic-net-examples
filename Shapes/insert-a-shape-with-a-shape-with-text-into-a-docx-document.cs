using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithText
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline text box shape (200 pt wide, 50 pt high) at the current cursor position.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);

        // The shape is already placed in the document, now add a paragraph and a run of text inside it.
        Paragraph shapeParagraph = textBox.FirstParagraph;               // Gets the first (and only) paragraph of the shape.
        Run run = new Run(doc) { Text = "Hello Aspose!" };               // Create a run with the desired text.
        shapeParagraph.AppendChild(run);                                 // Append the run to the shape's paragraph.

        // Optionally adjust alignment or other formatting of the shape here.
        // textBox.HorizontalAlignment = HorizontalAlignment.Center;
        // textBox.VerticalAlignment = VerticalAlignment.Top;

        // Save the document to a DOCX file.
        doc.Save("ShapeWithText.docx");
    }
}
