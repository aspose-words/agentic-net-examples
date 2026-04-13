using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ExtractVideoThumbnails
{
    public static void Main()
    {
        // Prepare deterministic file names.
        const string thumbnailImagePath = "thumb.png";
        const string documentPath = "sample.docx";
        const string outputFolder = "ExtractedThumbnails";

        // Ensure output folder exists.
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // Step 1: Create a sample PNG image to act as a video thumbnail.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 150;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle to make the image recognizable.
                g.DrawRectangle(new Pen(Aspose.Drawing.Color.Blue, 3), 10, 10, imgWidth - 20, imgHeight - 20);
            }
            // Save the bitmap as a PNG file.
            bitmap.Save(thumbnailImagePath);
        }

        // -------------------------------------------------
        // Step 2: Create a DOCX document and embed the thumbnail image.
        // In a real scenario this would be a video thumbnail, but for the demo we use a plain image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image as an inline shape.
        Shape thumbnailShape = builder.InsertImage(thumbnailImagePath);
        // Optionally set shape type to OLE object to mimic a video thumbnail (not required for extraction).
        // thumbnailShape.ShapeType = ShapeType.OleObject;
        doc.Save(documentPath);

        // -------------------------------------------------
        // Step 3: Load the document and extract all images (thumbnails) to PNG files.
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                // Force PNG extension if the original is not PNG.
                if (!extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
                {
                    extension = ".png";
                }

                string outFile = Path.Combine(outputFolder, $"thumbnail_{extractedCount}{extension}");
                // Save the image data to the file.
                shape.ImageData.Save(outFile);
                extractedCount++;
            }
        }

        // -------------------------------------------------
        // Validation: ensure at least one thumbnail was extracted.
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No thumbnail images were extracted from the document.");

        // Clean up temporary files (optional).
        File.Delete(thumbnailImagePath);
        File.Delete(documentPath);
    }
}
