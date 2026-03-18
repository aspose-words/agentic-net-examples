using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ResizePngImages
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Process only shapes that contain an image.
            if (!shape.HasImage) continue;

            // Get the image data and its size information.
            ImageData imgData = shape.ImageData;
            ImageSize imgSize = imgData.ImageSize;

            // Check if the image is a PNG by inspecting the first bytes (PNG signature: 89 50 4E 47).
            byte[] imageBytes = imgData.ImageBytes;
            if (imageBytes == null || imageBytes.Length < 4) continue;
            if (imageBytes[0] != 0x89 || imageBytes[1] != 0x50 || imageBytes[2] != 0x4E || imageBytes[3] != 0x47)
                continue; // Not a PNG image.

            // Original dimensions in pixels.
            int originalWidthPx = imgSize.WidthPixels;
            int originalHeightPx = imgSize.HeightPixels;

            // Desired height in pixels.
            const int targetHeightPx = 600;

            // Calculate scaling factor to preserve aspect ratio.
            double scale = (double)targetHeightPx / originalHeightPx;

            // Compute new width in pixels.
            int targetWidthPx = (int)Math.Round(originalWidthPx * scale);

            // Convert pixel dimensions to points (1 point = 1/72 inch).
            // Use the image's vertical resolution (DPI) for height conversion.
            double dpi = imgSize.VerticalResolution; // DPI (dots per inch)
            double heightPoints = targetHeightPx * 72.0 / dpi;
            double widthPoints = targetWidthPx * 72.0 / dpi;

            // Apply the new size to the shape. This preserves the aspect ratio.
            shape.Width = widthPoints;
            shape.Height = heightPoints;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
