using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Extract3DModelThumbnails
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG image that will act as a 3D model thumbnail.
        string thumbnailPath = Path.Combine(outputDir, "sample_thumbnail.png");
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple shape to make the image recognizable.
                using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    g.DrawEllipse(pen, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(thumbnailPath, ImageFormat.Png);
        }

        // Create a DOCX document and embed the thumbnail image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(thumbnailPath);
        string docPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docPath);

        // Load the document and extract all embedded images (thumbnails).
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine appropriate file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedPath = Path.Combine(outputDir, $"extracted_{extractedCount}{extension}");
                shape.ImageData.Save(extractedPath);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: indicate success.
        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to \"{outputDir}\".");
    }
}
