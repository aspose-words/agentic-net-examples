using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Step 1: Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Step 2: Insert two individual shapes that will later be grouped.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rectangle.Left = 50;   // Position relative to the page.
        rectangle.Top = 50;
        rectangle.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 300;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Step 3: Group the shapes using the DocumentBuilder.InsertGroupShape method.
        // This creates a GroupShape node and inserts it at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Optional: Adjust group properties (size, wrapping, etc.).
        group.Width = 500;
        group.Height = 300;
        group.WrapType = WrapType.None;

        // Step 4: Save the document as a DOCX file.
        doc.Save("GroupShape.docx");
    }
}
