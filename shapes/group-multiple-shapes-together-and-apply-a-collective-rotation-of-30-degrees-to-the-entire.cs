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
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rect.FillColor = Color.LightBlue;
        rect.Stroke.Color = Color.DarkBlue;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 120, 120);
        ellipse.FillColor = Color.LightCoral;
        ellipse.Stroke.Color = Color.DarkRed;

        // Group the two shapes together. The group bounds are calculated automatically.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Apply a collective rotation of 30 degrees to the entire group.
        group.Rotation = 30;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeRotation.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
