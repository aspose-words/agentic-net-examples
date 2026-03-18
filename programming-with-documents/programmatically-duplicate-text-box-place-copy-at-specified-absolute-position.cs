using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Insert the original floating text box.
        // ------------------------------------------------------------
        Shape originalTextBox = builder.InsertShape(ShapeType.TextBox, 150, 100);
        // Make the shape floating and position it relative to the page.
        originalTextBox.WrapType = WrapType.None;
        originalTextBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        originalTextBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        // Absolute position (in points) of the original text box.
        originalTextBox.Left = 100; // 100 points from the left edge of the page.
        originalTextBox.Top = 100;  // 100 points from the top edge of the page.

        // Add some content to the original text box.
        builder.MoveTo(originalTextBox.FirstParagraph);
        builder.Write("Original TextBox");

        // ------------------------------------------------------------
        // 2. Clone the text box.
        // ------------------------------------------------------------
        // The Clone method creates a deep copy of the shape (including its child nodes).
        Shape clonedTextBox = (Shape)originalTextBox.Clone(true);

        // ------------------------------------------------------------
        // 3. Re‑position the cloned text box at a new absolute location.
        // ------------------------------------------------------------
        clonedTextBox.Left = 300; // New X coordinate (points).
        clonedTextBox.Top = 200;  // New Y coordinate (points).

        // ------------------------------------------------------------
        // 4. Insert the cloned shape into the document.
        // ------------------------------------------------------------
        // InsertNode places the node at the current cursor position.
        // Move the builder to the end of the document body to append the clone.
        builder.MoveToDocumentEnd();
        builder.InsertNode(clonedTextBox);

        // ------------------------------------------------------------
        // 5. Save the resulting document.
        // ------------------------------------------------------------
        doc.Save("DuplicatedTextBox.docx");
    }
}
