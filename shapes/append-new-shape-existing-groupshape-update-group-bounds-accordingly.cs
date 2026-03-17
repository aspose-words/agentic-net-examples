using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AppendShapeToGroup
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder for inserting content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two initial shapes that will be grouped.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Blue;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 120, 120);
        shape2.Left = 250;
        shape2.Top = 40;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes. The group is inserted at the current builder position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Create a new shape that we want to add to the existing group.
        Shape newShape = new Shape(doc, ShapeType.Star);
        newShape.Width = 80;
        newShape.Height = 80;
        // Position the new shape relative to the group's internal coordinate system.
        // Here we place it partially outside the current bounds to demonstrate the update.
        newShape.Left = -30;   // left of the current origin
        newShape.Top = 180;    // below the current shapes
        newShape.FillColor = Color.Red;

        // Append the new shape to the group.
        group.AppendChild(newShape);

        // -----------------------------------------------------------------
        // Recalculate the group's bounds so that it fully contains all children.
        // -----------------------------------------------------------------
        // Determine the union of all child shapes' bounds in the group's local
        // coordinate space (the values returned by Shape.Bounds for children of a
        // GroupShape are expressed in that local space).
        float minLeft = 0, minTop = 0, maxRight = 0, maxBottom = 0;
        bool first = true;

        foreach (Shape child in group.GetChildNodes(NodeType.Shape, true))
        {
            RectangleF childBounds = child.Bounds; // local coordinates

            if (first)
            {
                minLeft = childBounds.Left;
                minTop = childBounds.Top;
                maxRight = childBounds.Right;
                maxBottom = childBounds.Bottom;
                first = false;
            }
            else
            {
                if (childBounds.Left < minLeft) minLeft = childBounds.Left;
                if (childBounds.Top < minTop) minTop = childBounds.Top;
                if (childBounds.Right > maxRight) maxRight = childBounds.Right;
                if (childBounds.Bottom > maxBottom) maxBottom = childBounds.Bottom;
            }
        }

        // Adjust the group's outer bounds (document coordinates) and its internal
        // coordinate system (origin and size) to match the new union.
        // The group's current outer bounds.
        RectangleF oldGroupBounds = group.Bounds;

        // New outer bounds: shift by the difference between the old origin (0,0)
        // in the group's local space and the new minimum coordinates.
        group.Bounds = new RectangleF(
            oldGroupBounds.Left + minLeft,
            oldGroupBounds.Top + minTop,
            maxRight - minLeft,
            maxBottom - minTop);

        // Update the internal coordinate system so that (0,0) corresponds to the
        // new top‑left corner of the group's content.
        group.CoordOrigin = new Point((int)minLeft, (int)minTop);
        group.CoordSize = new Size((int)(maxRight - minLeft), (int)(maxBottom - minTop));

        // Save the resulting document.
        doc.Save("AppendShapeToGroup.docx");
    }
}
