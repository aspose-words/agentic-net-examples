using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

class InsertShapesExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle shape (width: 150 points, height: 100 points).
        Shape inlineRect = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        // Set a red border for the inline shape.
        inlineRect.Stroke.Color = Color.Red;

        // Insert a floating ellipse shape.
        // Position: 100 points from the left and top of the page.
        // Size: 200x150 points. No text wrapping.
        Shape floatingEllipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            200, 150,
            WrapType.None);
        // Set a blue fill for the floating shape.
        floatingEllipse.Fill.ForeColor = Color.LightBlue;

        // Insert a text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 250, 80);
        textBox.WrapType = WrapType.Square;
        textBox.Fill.ForeColor = Color.LightYellow;
        // Add a paragraph with text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello Aspose.Words!");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Save the document as a DOCM file (macro-enabled Word document).
        // No special compliance options are required for these primitive shapes.
        doc.Save("ShapesInserted.docm", SaveFormat.Docm);
    }
}
