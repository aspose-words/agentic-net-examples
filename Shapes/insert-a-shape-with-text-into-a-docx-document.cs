using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeWithText
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline text box shape with the desired size.
        // The InsertShape method returns the created Shape object.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Move the builder's cursor inside the text box (to its last paragraph).
        builder.MoveTo(textBoxShape.LastParagraph);

        // Write the text that will appear inside the shape.
        builder.Write("Hello world!");

        // Save the document to a DOCX file.
        doc.Save("ShapeWithText.docx");
    }
}
