using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will be grouped.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rect.Left = 50;
        rect.Top = 50;
        rect.FillColor = System.Drawing.Color.LightBlue;
        rect.Stroke.Color = System.Drawing.Color.DarkBlue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 120, 120);
        ellipse.Left = 250;
        ellipse.Top = 80;
        ellipse.FillColor = System.Drawing.Color.LightCoral;
        ellipse.Stroke.Color = System.Drawing.Color.DarkRed;

        // Group the two shapes together. The builder will calculate the group bounds automatically.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Apply a collective rotation of 30 degrees to the whole group.
        group.Rotation = 30;

        // Save the document.
        string outputPath = "GroupShapeRotation.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
