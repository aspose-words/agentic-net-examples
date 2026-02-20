using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace InsertShapeWithText
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a floating textbox shape (width: 200 points, height: 100 points).
            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

            // Optional: configure the textbox to grow automatically to fit its contents.
            textBoxShape.TextBox.FitShapeToText = true;

            // Move the cursor inside the textbox's last paragraph so we can write text there.
            builder.MoveTo(textBoxShape.LastParagraph);

            // Set desired font properties for the text inside the shape.
            builder.Font.Size = 14;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;

            // Write the text that will appear inside the shape.
            builder.Write("Hello, Aspose.Words!");

            // Save the document to a DOCX file.
            doc.Save("ShapeWithText.docx");
        }
    }
}
