using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph so the document contains some text.
        builder.Writeln("Document with a group shape.");

        // Insert the first shape (a rectangle) and set its position and outline color.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;
        rect.Top = 50;
        rect.Stroke.Color = Color.Blue;

        // Insert the second shape (an ellipse) and set its position and outline color.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 120;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The builder automatically calculates the group's position and size.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: adjust group properties (e.g., make it float behind text).
        group.WrapType = WrapType.None;
        group.BehindText = true;

        // Save the document as a plain‑text file.
        doc.Save("GroupShapeInTxt.txt", SaveFormat.Text);
    }
}
