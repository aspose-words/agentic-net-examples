using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ResizeBmpImages
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Input.docx";

        // Folder where the processed document will be saved.
        string outputPath = @"C:\Docs\Output.docx";

        // Load the document.
        Document doc = new Document(inputPath);

        // Desired pixel dimensions.
        const int targetWidthPixels = 640;
        const int targetHeightPixels = 480;

        // Desired DPI for conversion (default Aspose.Words DPI is 96).
        const double dpi = 96.0;

        // Convert pixel dimensions to points (1 point = 1/72 inch).
        // points = pixels * 72 / dpi
        double targetWidthPoints = targetWidthPixels * 72.0 / dpi;
        double targetHeightPoints = targetHeightPixels * 72.0 / dpi;

        // Iterate through all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Process only shapes that contain an image.
            if (!shape.HasImage)
                continue;

            // Use Aspose.Words' ImageType to check for BMP without System.Drawing.
            if (shape.ImageData.ImageType != ImageType.Bmp)
                continue;

            // Resize the shape to the target dimensions.
            // This changes the displayed size; the underlying bitmap remains unchanged,
            // which is sufficient for legacy applications that rely on shape size.
            shape.Width = targetWidthPoints;
            shape.Height = targetHeightPoints;
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
