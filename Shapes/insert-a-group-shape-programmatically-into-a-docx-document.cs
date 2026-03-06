using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

namespace GroupShapeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder to work with the document.
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

            // Group the two shapes into a new GroupShape node.
            // The InsertGroupShape method automatically calculates the position and size of the group.
            GroupShape group = builder.InsertGroupShape(shape1, shape2);

            // Optionally, add another shape to the group by cloning an existing shape.
            Shape shape3 = (Shape)shape1.Clone(true);
            GroupShape groupWithClone = builder.InsertGroupShape(group, shape3);

            // Save the document to a DOCX file.
            doc.Save("GroupShapeExample.docx");
        }
    }
}
