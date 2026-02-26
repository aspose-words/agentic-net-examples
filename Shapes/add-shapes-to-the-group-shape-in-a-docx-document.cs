using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroup
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert individual shapes.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rectangle.Left = 20;   // Position within the page.
        rectangle.Top = 20;
        rectangle.Stroke.Color = Color.Blue;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 250;
        ellipse.Top = 30;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The InsertGroupShape method automatically calculates the
        // position and size of the new GroupShape node and inserts it at the current builder position.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Add a third shape (a star) directly to the existing group.
        Shape star = new Shape(doc, ShapeType.Star)
        {
            Width = 80,
            Height = 80,
            // Position the star relative to the group's internal coordinate system.
            // Here we place it near the centre of the group.
            Left = (group.Bounds.Width - 80) / 2,
            Top = (group.Bounds.Height - 80) / 2,
            FillColor = Color.Red
        };
        group.AppendChild(star);

        // Save the document to a DOCX file.
        doc.Save("AddShapesToGroup.docx");
    }
}
