using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color, Pen, Rectangle

public class Extract3DModelThumbnails
{
    public static void Main()
    {
        // Define deterministic file and folder names.
        const string workDir = "Artifacts";
        const string thumbnailFileName = "thumbnail.png";
        const string docFileName = "Sample.docx";

        // Ensure the output directory exists.
        Directory.CreateDirectory(workDir);

        // -------------------------------------------------
        // 1. Create a sample thumbnail image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        string thumbnailPath = Path.Combine(workDir, thumbnailFileName);

        using (var bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);

                // Draw a recognizable rectangle.
                var rect = new Rectangle(20, 20, imgWidth - 40, imgHeight - 40);
                using (var pen = new Pen(Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, rect);
                }
            }

            // Save the bitmap as PNG – this will act as the 3‑D model thumbnail.
            bitmap.Save(thumbnailPath);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and insert the thumbnail image.
        // -------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Document containing a 3‑D model thumbnail:");
        // Insert the image as an inline shape.
        builder.InsertImage(thumbnailPath);
        // Save the document.
        string docPath = Path.Combine(workDir, docFileName);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document (simulating a user‑provided file) and extract all images.
        // -------------------------------------------------
        var loadedDoc = new Document(docPath);
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine a proper file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                // Force PNG extension if the original type is not PNG.
                if (!extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
                {
                    extension = ".png";
                }

                string outFileName = Path.Combine(workDir, $"ExtractedThumbnail_{imageIndex}{extension}");
                // Save the image data to the file system.
                shape.ImageData.Save(outFileName);
                extractedCount++;
                imageIndex++;
            }
        }

        // -------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // -------------------------------------------------
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional: verify that the extracted files exist.
        for (int i = 0; i < imageIndex; i++)
        {
            string expectedPath = Path.Combine(workDir, $"ExtractedThumbnail_{i}.png");
            if (!File.Exists(expectedPath))
                throw new FileNotFoundException($"Expected extracted image not found: {expectedPath}");
        }

        // The example finishes without requiring user interaction.
    }
}
