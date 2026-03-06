using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will later be grouped.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.Page, 100,
                                         RelativeVerticalPosition.Page, 100, 200, 250, WrapType.None);
        rect.Stroke.Color = Color.Red;   // Red border for the rectangle.

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, RelativeHorizontalPosition.Page, 150,
                                            RelativeVerticalPosition.Page, 150, 150, 200, WrapType.None);
        ellipse.Stroke.Color = Color.Green; // Green border for the ellipse.

        // Group the two shapes. The builder will calculate the group position and size automatically.
        GroupShape autoGroup = builder.InsertGroupShape(rect, ellipse);

        // Optionally, create another group with explicit size and position.
        double left = 300;   // Distance from the page origin to the left side of the group.
        double top = 300;    // Distance from the page origin to the top side of the group.
        double width = 250;  // Width of the group.
        double height = 250; // Height of the group.

        // Clone one of the existing shapes to reuse it inside the new group.
        Shape rectClone = (Shape)rect.Clone(true);

        // Insert a group shape with specified dimensions.
        GroupShape sizedGroup = builder.InsertGroupShape(left, top, width, height, rectClone, ellipse);

        // Save the document to a DOCX file.
        string artifactsDir = "output/";
        System.IO.Directory.CreateDirectory(artifactsDir);
        doc.Save(System.IO.Path.Combine(artifactsDir, "GroupShapeExample.docx"));
    }
}
