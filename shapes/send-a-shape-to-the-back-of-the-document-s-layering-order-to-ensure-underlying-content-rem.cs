using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first floating rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 200);
        rectangle.FillColor = Color.LightBlue;
        rectangle.Left = 50;   // Position from the left edge of the page.
        rectangle.Top = 50;    // Position from the top edge of the page.
        rectangle.WrapType = WrapType.None; // Make it a floating shape.

        // Insert the second floating ellipse shape that overlaps the rectangle.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 200, 200);
        ellipse.FillColor = Color.LightCoral;
        ellipse.Left = 100;
        ellipse.Top = 100;
        ellipse.WrapType = WrapType.None;

        // Adjust ZOrder so that the rectangle is behind the ellipse.
        // Lower ZOrder values are rendered behind higher values.
        rectangle.ZOrder = 0;
        ellipse.ZOrder = 1;

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeBack.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Optionally, inform that the process completed (no interactive prompts required).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
