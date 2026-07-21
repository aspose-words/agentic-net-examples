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

        // Insert a single‑corner snipped rectangle shape.
        // Width = 200 points, Height = 100 points.
        Shape snipShape = builder.InsertShape(ShapeType.SingleCornerSnipped, 200, 100);

        // Set a light gray fill color.
        snipShape.FillColor = Color.LightGray;

        // The Adjustments collection is read‑only in this API version.
        // The default adjustment gives a visible snip, so we skip explicit assignment.

        // Position the shape as a floating object.
        snipShape.WrapType = WrapType.None;
        snipShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        snipShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        snipShape.Left = 100; // points from the left edge of the page
        snipShape.Top = 100;  // points from the top edge of the page

        // Save the document.
        string outputPath = "SnipCornerShape.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
