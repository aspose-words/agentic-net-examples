using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 100, 100);
        rectangle.Left = 50;   // Position from the left edge of the page.
        rectangle.Top = 50;    // Position from the top edge of the page.
        rectangle.Stroke.Color = Color.Blue;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 80, 80);
        ellipse.Left = 200;
        ellipse.Top = 70;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes together. The group is inserted at the builder's current position.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Apply a collective rotation of 30 degrees to the entire group.
        group.Rotation = 30;

        // Save the document to disk.
        const string outputPath = "GroupRotation.docx";
        doc.Save(outputPath);

        // Validate that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
