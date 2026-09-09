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

            // Insert the original floating text box.
            Shape originalBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
            originalBox.WrapType = WrapType.None;
            originalBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            originalBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            originalBox.Left = 50;   // Position of the original box.
            originalBox.Top = 50;

            // Add some text to the original text box.
            builder.MoveTo(originalBox.FirstParagraph);
            builder.Write("Original TextBox");

            // Clone the original shape (deep copy).
            Shape clonedBox = (Shape)originalBox.Clone(true);

            // Set the absolute position for the cloned text box.
            clonedBox.Left = 300;   // Horizontal position in points.
            clonedBox.Top = 200;    // Vertical position in points.
            clonedBox.WrapType = WrapType.None;
            clonedBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            clonedBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Add some text to the cloned text box.
            builder.MoveTo(clonedBox.FirstParagraph);
            builder.Write("Cloned TextBox");

            // Insert the cloned shape into the document.
            doc.FirstSection.Body.FirstParagraph.AppendChild(clonedBox);

            // Save the document.
            doc.Save("DuplicatedTextBox.docx");
        }
    }
}
