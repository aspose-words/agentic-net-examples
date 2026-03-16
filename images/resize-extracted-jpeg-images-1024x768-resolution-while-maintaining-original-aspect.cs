using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class ResizeExtractedJpegImages
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Maximum dimensions in pixels.
        const int maxWidthPixels = 1024;
        const int maxHeightPixels = 768;

        // Iterate over all shapes that contain images.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Skip shapes that do not contain an image.
            if (!shape.HasImage) continue;

            // Get the original image size in pixels.
            ImageSize imgSize = shape.ImageData.ImageSize;
            int origWidthPx = imgSize.WidthPixels;
            int origHeightPx = imgSize.HeightPixels;

            // Determine the scaling factor that fits the image inside the max bounds
            // while preserving the original aspect ratio.
            double widthRatio = (double)maxWidthPixels / origWidthPx;
            double heightRatio = (double)maxHeightPixels / origHeightPx;
            double scale = Math.Min(1.0, Math.Min(widthRatio, heightRatio)); // never upscale

            // Compute the new size in pixels.
            int newWidthPx = (int)Math.Round(origWidthPx * scale);
            int newHeightPx = (int)Math.Round(origHeightPx * scale);

            // Convert pixel dimensions to points (1 point = 1/72 inch).
            // Points = pixels * 72 / DPI.
            double dpiX = imgSize.HorizontalResolution;
            double dpiY = imgSize.VerticalResolution;
            double newWidthPt = newWidthPx * 72.0 / dpiX;
            double newHeightPt = newHeightPx * 72.0 / dpiY;

            // Apply the new dimensions to the shape.
            shape.Width = newWidthPt;
            shape.Height = newHeightPt;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
