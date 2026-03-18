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

        // Path to the file that will be embedded as an OLE object.
        // Adjust the path to point to an existing file on your system.
        string oleFilePath = @"C:\Temp\Sample.xlsx";

        // Insert the OLE object into the document.
        // Parameters: file name, isLinked (false = embed), asIcon (false = display content), presentation (null = default image).
        Shape oleShape = builder.InsertOleObject(oleFilePath, false, false, null);

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

        // Store the dimensions for later layout calculations.
        // Here we simply output them to the console, but they could be saved to a data structure.
        Console.WriteLine($"OLE object display size: {displayWidthPoints:F2} pt × {displayHeightPoints:F2} pt");
        Console.WriteLine($"Converted to pixels (96 DPI): {widthPixels} px × {heightPixels} px");

        // Save the document.
        string outputPath = @"C:\Temp\OleObjectWithDimensions.docx";
        doc.Save(outputPath);
    }
}
