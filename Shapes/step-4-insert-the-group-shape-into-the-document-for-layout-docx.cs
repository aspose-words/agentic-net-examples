using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) and set its position and stroke color.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        // Insert the second shape (an ellipse) and set its position and stroke color.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes. The InsertGroupShape method creates a GroupShape node,
        // automatically calculates its position and size, and inserts it at the current cursor location.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Example: set the group's wrap type to None so it behaves as a floating object.
        group.WrapType = WrapType.None;

        // Save the document in DOCX format.
        doc.Save("GroupShapeLayout.docx");
    }
}
