using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

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
        rectangle.FillColor = Color.LightBlue;      // Set fill color.
        rectangle.StrokeColor = Color.DarkBlue;     // Set outline color.

        // Insert a floating text box shape.
        Shape textBox = builder.InsertShape(
            ShapeType.TextBox,
            RelativeHorizontalPosition.Page, 100,   // Left position.
            RelativeVerticalPosition.Page, 150,     // Top position.
            200,                                     // Width.
            100,                                     // Height.
            WrapType.None);                          // No text wrapping.

        // Add a paragraph with text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello Aspose!");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Save the document as a DOCX file.
        doc.Save("Shapes.docx");
    }
}
