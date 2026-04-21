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

        // Insert the first floating shape (rectangle) and set its size and position.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 100, 60);
        rect.Left = 50;
        rect.Top = 50;
        rect.Stroke.Color = System.Drawing.Color.Blue;

        // Insert the second floating shape (ellipse) and set its size and position.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 80, 80);
        ellipse.Left = 180;
        ellipse.Top = 120;
        ellipse.Stroke.Color = System.Drawing.Color.Green;

        // Group the two shapes together. The group is inserted at the current builder position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Apply a collective rotation of 30 degrees to the entire group.
        group.Rotation = 30;

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GroupShapeRotation.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Optionally, inform that the process completed successfully.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
