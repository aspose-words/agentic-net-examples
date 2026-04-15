using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsDuplicateTextBox
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a floating text box shape.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
            textBox.WrapType = WrapType.None;
            // Position the original text box at (50, 50) points from the top‑left corner of the page.
            textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            textBox.Left = 50;
            textBox.Top = 50;

            // Add some text inside the original text box.
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Original Text Box");

            // Clone the text box node (deep copy).
            Shape clonedBox = (Shape)textBox.Clone(true);

            // Set the cloned box to a new absolute position, e.g., (250, 150) points.
            clonedBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            clonedBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            clonedBox.Left = 250;
            clonedBox.Top = 150;

            // Insert the cloned text box into the document.
            // Here we add it after the original shape in the body.
            doc.FirstSection.Body.FirstParagraph.AppendChild(clonedBox);

            // Save the resulting document.
            string outputPath = "DuplicatedTextBox.docx";
            doc.Save(outputPath);
        }
    }
}
