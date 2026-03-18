using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

class ApplyBorderToExtractedPngImages
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("MyDir/Images.docx");

        // Get all shape nodes (they may contain images).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Apply a 5‑point red border to each side of the inline image.
            foreach (Border border in shape.ImageData.Borders)
            {
                border.Color = Color.Red;          // Red color.
                border.LineWidth = 5.0;            // Width in points (approx. 5 px).
                border.LineStyle = LineStyle.Single; // Ensure the border is visible.
            }

            // Build a file name based on the image index and its original extension.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string fileName = $"File.ExtractImages.{imageIndex}{extension}";

            // Save the image (with the applied border) to the file system.
            shape.ImageData.Save(Path.Combine("ArtifactsDir", fileName));

            imageIndex++;
        }
    }
}
