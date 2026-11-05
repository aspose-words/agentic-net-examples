using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables; // Required for NodeType

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an empty paragraph to host the group shape.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a GroupShape and set its initial size and positioning.
        GroupShape group = new GroupShape(doc)
        {
            Width = 300,
            Height = 200,
            WrapType = WrapType.None,
            RelativeHorizontalPosition = RelativeHorizontalPosition.Margin,
            RelativeVerticalPosition = RelativeVerticalPosition.Margin,
            Left = 0,
            Top = 0
        };

        // Append the group shape to the paragraph.
        paragraph.AppendChild(group);

        // Add the first shape (a rectangle) to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 50,
            Left = 0,
            Top = 0,
            FillColor = System.Drawing.Color.LightBlue,
            StrokeColor = System.Drawing.Color.DarkBlue
        };
        group.AppendChild(rect);

        // Update the group bounds after adding the first shape.
        UpdateGroupShapeBounds(group);

        // Append a new shape (an ellipse) to the existing group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 80,
            Height = 80,
            Left = 120,
            Top = 30,
            FillColor = System.Drawing.Color.LightCoral,
            StrokeColor = System.Drawing.Color.Maroon
        };
        group.AppendChild(ellipse);

        // Update the group bounds to encompass the newly added shape.
        UpdateGroupShapeBounds(group);

        // Save the document.
        string outputPath = "GroupShapeExample.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
    }

    // Calculates the bounding rectangle of all child shapes and updates the group's size.
    private static void UpdateGroupShapeBounds(GroupShape group)
    {
        double maxRight = 0;
        double maxBottom = 0;

        foreach (Shape child in group.GetChildNodes(NodeType.Shape, true))
        {
            double right = child.Left + child.Width;
            double bottom = child.Top + child.Height;

            if (right > maxRight) maxRight = right;
            if (bottom > maxBottom) maxBottom = bottom;
        }

        group.Width = maxRight;
        group.Height = maxBottom;
    }
}
