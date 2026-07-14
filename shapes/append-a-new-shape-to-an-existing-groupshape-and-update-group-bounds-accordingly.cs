using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two initial shapes.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 100, 80);
        rect.Left = 20;
        rect.Top = 20;
        rect.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 120, 60);
        ellipse.Left = 150;
        ellipse.Top = 30;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. InsertGroupShape calculates the initial bounds automatically.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Create a new shape using DocumentBuilder.InsertShape so that its markup language matches the group (DML).
        Shape star = builder.InsertShape(ShapeType.Star, 70, 70);
        star.Left = 100;   // Position relative to the group's coordinate system.
        star.Top = 100;
        star.FillColor = Color.Red;

        // Append the new shape to the existing group.
        group.AppendChild(star);

        // Recalculate the group's bounds to include all child shapes.
        float minLeft = float.MaxValue;
        float minTop = float.MaxValue;
        float maxRight = float.MinValue;
        float maxBottom = float.MinValue;

        foreach (Shape child in group.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            // Skip shapes with no size.
            if (child.Width <= 0 || child.Height <= 0) continue;

            float left = (float)child.Left;
            float top = (float)child.Top;
            float right = left + (float)child.Width;
            float bottom = top + (float)child.Height;

            if (left < minLeft) minLeft = left;
            if (top < minTop) minTop = top;
            if (right > maxRight) maxRight = right;
            if (bottom > maxBottom) maxBottom = bottom;
        }

        // Apply the new bounds to the group shape.
        group.Bounds = new RectangleF(minLeft, minTop, maxRight - minLeft, maxBottom - minTop);

        // Save the document.
        string outputPath = "GroupShapeAppendShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
