using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ResizeExtractedImages
{
    static void Main()
    {
        // Load the source document that contains images.
        Document doc = new Document("Input.docx");

        // Desired square size in pixels.
        const int targetSizePixels = 500;

        // Convert pixels to points (1 point = 1/72 inch, default DPI = 96).
        // Points = pixels * 72 / DPI
        double targetSizePoints = targetSizePixels * 72.0 / 96.0;

        // Iterate over all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Set both width and height to the target square size (in points).
                shape.Width = targetSizePoints;
                shape.Height = targetSizePoints;

                // Fit the image to the shape while preserving its aspect ratio.
                // This adds transparent padding where the aspect ratios differ.
                shape.ImageData.FitImageToShape();
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
