using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace DuplicateTextBoxExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a floating text box shape.
            Shape originalBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
            originalBox.WrapType = WrapType.None;
            originalBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            originalBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            originalBox.Left = 100;   // X position in points.
            originalBox.Top = 100;    // Y position in points.

            // Add some text to the original text box.
            builder.MoveTo(originalBox.FirstParagraph);
            builder.Write("Original TextBox");

            // Clone the text box (deep copy of the shape and its contents).
            Shape clonedBox = (Shape)originalBox.Clone(true);
            // Position the cloned box at a different absolute location.
            clonedBox.Left = 300;   // New X position.
            clonedBox.Top = 200;    // New Y position.

            // Insert the cloned shape into the document.
            // Here we add it after the original shape in the body.
            doc.FirstSection.Body.FirstParagraph.AppendChild(clonedBox);

            // Save the document to a file.
            doc.Save("DuplicatedTextBox.docx");
        }
    }
}
