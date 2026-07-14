using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample thumbnail image (PNG) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string thumbPath = Path.Combine(outputDir, "thumb.png");
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightBlue);
        // Dispose drawing objects.
        graphics.Dispose();
        bitmap.Save(thumbPath);
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Create a DOCX document and insert the thumbnail image.
        //    In a real scenario this would be the video thumbnail.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(thumbPath);
        string docPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images from shapes.
        //    These represent the embedded video thumbnails.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                // Force PNG extension if the image is not already PNG.
                if (!extension.Equals(".png", StringComparison.OrdinalIgnoreCase))
                    extension = ".png";

                string imageFileName = $"VideoThumbnail_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputDir, imageFileName);
                shape.ImageData.Save(imagePath);
                imageIndex++;
            }
        }

        // Validate that at least one thumbnail was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No video thumbnail images were extracted.");
    }
}
