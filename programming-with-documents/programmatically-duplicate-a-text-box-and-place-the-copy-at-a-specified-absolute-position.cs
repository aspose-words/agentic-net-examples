using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the original floating text box.
        // Width = 200 points, Height = 50 points.
        Shape originalBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        originalBox.WrapType = WrapType.None;
        originalBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        originalBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        originalBox.Left = 100;   // 100 points from the left edge of the page.
        originalBox.Top = 100;    // 100 points from the top edge of the page.

        // Add some text to the original text box.
        builder.MoveTo(originalBox.FirstParagraph);
        builder.Write("Original TextBox");

        // Clone the text box node (deep clone).
        Shape clonedBox = (Shape)originalBox.Clone(true);

        // Position the cloned box at a different absolute location.
        clonedBox.Left = 300;   // 300 points from the left edge of the page.
        clonedBox.Top = 300;    // 300 points from the top edge of the page.

        // Insert the cloned box into the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(clonedBox);

        // Save the document.
        doc.Save("DuplicatedTextBox.docx");
    }
}
