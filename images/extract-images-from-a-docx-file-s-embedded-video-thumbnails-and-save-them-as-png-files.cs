using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class ExtractVideoThumbnails
{
    public static void Main()
    {
        // Folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample thumbnail image (PNG) that will act as a video thumbnail.
        // -----------------------------------------------------------------
        string thumbnailPath = Path.Combine(artifactsDir, "thumb.png");
        using (var bitmap = new Bitmap(200, 200))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Simple deterministic drawing – a filled rectangle.
                graphics.FillRectangle(new SolidBrush(Color.DarkBlue), 50, 50, 100, 100);
            }
            bitmap.Save(thumbnailPath);
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCX document and insert the thumbnail image.
        //    In a real scenario this would be the thumbnail of an embedded video.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(thumbnailPath);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images (thumbnails) from shapes.
        // -----------------------------------------------------------------
        var loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine proper file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(artifactsDir, $"extracted_{extractedCount}{extension}");

                // Save the image to the file system.
                shape.ImageData.Save(outFile);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No thumbnail images were extracted from the document.");
    }
}
