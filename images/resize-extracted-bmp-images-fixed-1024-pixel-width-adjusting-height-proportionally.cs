using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ResizeBmpImages
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            // Process only shapes that contain an image and the image is a BMP.
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Bmp)
            {
                // Get original image dimensions in pixels.
                ImageSize imgSize = shape.ImageData.ImageSize;
                int originalWidthPx = imgSize.WidthPixels;
                int originalHeightPx = imgSize.HeightPixels;

                // Use the image's horizontal DPI for conversion (vertical DPI is the same for BMPs).
                double dpi = imgSize.HorizontalResolution;

                // Desired width in pixels.
                const int targetWidthPx = 1024;

                // Calculate scaling factor and the new height to keep the aspect ratio.
                double scale = (double)targetWidthPx / originalWidthPx;
                int targetHeightPx = (int)Math.Round(originalHeightPx * scale);

                // Convert pixel dimensions to points (1 point = 1/72 inch).
                double targetWidthPt = targetWidthPx * 72.0 / dpi;
                double targetHeightPt = targetHeightPx * 72.0 / dpi;

                // Apply the new size to the shape.
                shape.Width = targetWidthPt;
                shape.Height = targetHeightPt;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
