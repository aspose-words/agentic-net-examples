using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoTxt
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will be grouped.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        shape1.Left = 50;   // Position from the left edge of the page.
        shape1.Top = 50;    // Position from the top edge of the page.
        shape1.Stroke.Color = Color.Blue;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 300;
        shape2.Top = 100;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape is inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Optionally set properties on the group shape (e.g., make it non‑overlapping).
        group.AllowOverlap = false;

        // Save the document as a plain‑text file. The group shape will not be visible in the TXT output,
        // but the operation demonstrates that the group shape exists in the document structure.
        doc.Save("GroupShapeDocument.txt", SaveFormat.Text);
    }
}
