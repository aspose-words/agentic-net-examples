using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class AdjustOleIconDimensions
{
    public static void Main()
    {
        // Prepare a temporary directory for sample files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Create a simple text file that will be embedded as an OLE object.
        string oleFilePath = Path.Combine(dataDir, "sample.txt");
        File.WriteAllText(oleFilePath, "This is a sample OLE object content.");

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object as an icon. No custom icon file or caption is provided,
        // so Aspose.Words will use a predefined icon.
        Shape oleShape = builder.InsertOleObjectAsIcon(oleFilePath, false, null, null);

        // Lock the aspect ratio to ensure proportional scaling.
        oleShape.AspectRatioLocked = true;

        // Retrieve the original icon image size.
        ImageSize originalSize = oleShape.ImageData.ImageSize;
        double aspectRatio = originalSize.WidthPoints / originalSize.HeightPoints;

        // Define a new width (in points) and calculate the corresponding height
        // to maintain the original aspect ratio.
        double newWidth = 100.0; // points
        double newHeight = newWidth / aspectRatio;

        // Apply the new dimensions to the OLE icon shape.
        oleShape.Width = newWidth;
        oleShape.Height = newHeight;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AdjustedOleIcon.docx");
        doc.Save(outputPath);
    }
}
