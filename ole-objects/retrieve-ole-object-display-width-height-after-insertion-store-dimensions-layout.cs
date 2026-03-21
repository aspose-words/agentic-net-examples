using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;

class OleObjectDimensions
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a temporary text file to embed as an OLE object.
        string tempFilePath = Path.Combine(Path.GetTempPath(), "Sample.txt");
        File.WriteAllText(tempFilePath, "Sample content for OLE object.");

        // Insert the OLE object into the document.
        // Parameters: file name, isLinked (false = embed), asIcon (false = display content), presentation (null = default image).
        Shape oleShape = builder.InsertOleObject(tempFilePath, false, false, null);

        // After insertion the shape's Width and Height properties contain the display size in points.
        double displayWidthPoints = oleShape.Width;   // Width in points (1 point = 1/72 inch)
        double displayHeightPoints = oleShape.Height; // Height in points

        // Convert the size to pixels (e.g., 96 DPI) using ShapeRenderer.
        // ShapeRenderer works for any Shape, including OLE objects.
        ShapeRenderer renderer = new ShapeRenderer(oleShape);
        // Scale = 1.0 (100%), DPI = 96 (standard screen resolution)
        Size sizeInPixels = renderer.GetSizeInPixels(1.0f, 96.0f);
        int widthPixels = sizeInPixels.Width;
        int heightPixels = sizeInPixels.Height;

        // Output the dimensions.
        Console.WriteLine($"OLE object display size: {displayWidthPoints:F2} pt × {displayHeightPoints:F2} pt");
        Console.WriteLine($"Converted to pixels (96 DPI): {widthPixels} px × {heightPixels} px");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OleObjectWithDimensions.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");

        // Clean up the temporary file.
        File.Delete(tempFilePath);
    }
}
