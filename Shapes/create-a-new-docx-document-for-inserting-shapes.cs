using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle shape (width: 100 points, height: 50 points).
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Set a light blue fill color for the rectangle.
        rectangle.FillColor = Color.LightBlue;

        // Insert a floating text box shape.
        // Position: 150 points from the left of the page, 200 points from the top of the page.
        // Size: 200 points wide, 100 points high. No text wrapping.
        Shape textBox = builder.InsertShape(
            ShapeType.TextBox,
            RelativeHorizontalPosition.Page, 150,
            RelativeVerticalPosition.Page, 200,
            200, 100,
            WrapType.None);

        // Add a paragraph with a run of text inside the text box.
        Paragraph tbParagraph = new Paragraph(doc);
        Run tbRun = new Run(doc, "Hello Aspose!");
        tbParagraph.AppendChild(tbRun);
        textBox.AppendChild(tbParagraph);

        // Save the document as a DOCX file.
        doc.Save("Shapes.docx");
    }
}
