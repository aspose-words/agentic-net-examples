using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape as a floating shape.
        Shape shape1 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left
            RelativeVerticalPosition.Page, 100,     // top
            200, 250,                               // width, height
            WrapType.None);
        shape1.Stroke.Color = Color.Red;

        // Insert an ellipse shape as a floating shape.
        Shape shape2 = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 150,   // left
            RelativeVerticalPosition.Page, 150,     // top
            150, 200,                               // width, height
            WrapType.None);
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes. The group’s position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Clone one of the shapes and add it to the existing group.
        Shape shape3 = (Shape)shape1.Clone(true);
        shape3.Stroke.Color = Color.Blue;
        // Insert a new group that contains the previous group and the cloned shape.
        builder.InsertGroupShape(group, shape3);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx");
    }
}
