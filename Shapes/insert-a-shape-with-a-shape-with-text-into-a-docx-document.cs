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

        // Insert a textbox shape with specific dimensions (300x100 points).
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 300, 100);

        // Access the TextBox object to configure its properties if needed.
        TextBox textBox = textBoxShape.TextBox;
        // Example: make the shape grow to fit the text.
        textBox.FitShapeToText = true;
        // Example: set internal margins (optional).
        textBox.InternalMarginTop = 5;
        textBox.InternalMarginBottom = 5;
        textBox.InternalMarginLeft = 5;
        textBox.InternalMarginRight = 5;

        // Move the builder cursor to the last paragraph inside the textbox shape.
        builder.MoveTo(textBoxShape.LastParagraph);

        // Write the desired text into the textbox.
        builder.Font.Size = 14;
        builder.Font.Name = "Arial";
        builder.Write("This is a textbox shape with text inserted using Aspose.Words.");

        // Save the document to a DOCX file.
        doc.Save("ShapeWithText.docx");
    }
}
