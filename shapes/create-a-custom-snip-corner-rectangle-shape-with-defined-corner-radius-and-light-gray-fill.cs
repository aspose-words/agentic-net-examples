using System;
using System.Drawing;
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

        // Insert a rectangle shape (used as a snip‑corner rectangle) with specific size.
        Shape snipRect = builder.InsertShape(ShapeType.Rectangle, 200, 100);

        // Apply a light gray fill.
        snipRect.FillColor = Color.LightGray;

        // Optional: set a visible border.
        snipRect.StrokeColor = Color.Black;
        snipRect.StrokeWeight = 1.0; // line width in points

        // Save the document.
        string outputPath = "CustomSnipCornerRectangle.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
