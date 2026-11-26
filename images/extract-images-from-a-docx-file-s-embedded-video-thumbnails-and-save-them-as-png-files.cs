using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Deterministic file names
        const string thumbnailPath = "thumb.png";
        const string docPath = "sample.docx";

        // -------------------------------------------------
        // Create a sample thumbnail image (200x200, LightBlue)
        // -------------------------------------------------
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
            }
            bitmap.Save(thumbnailPath);
        }

        // -------------------------------------------------
        // Create a document and insert the thumbnail as an image shape
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an image shape and set its image data to the thumbnail file
        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.ImageData.SetImage(thumbnailPath);

        // Insert the shape into the document (inside a paragraph)
        builder.InsertNode(imageShape);
        doc.Save(docPath);

        // -------------------------------------------------
        // Load the document and extract all image shapes
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Ensure the shape actually contains an image
            if (shape.HasImage)
            {
                string outputImage = $"thumbnail-{extractedCount}.png";
                shape.ImageData.Save(outputImage);
                if (!File.Exists(outputImage))
                    throw new Exception($"Failed to save extracted image: {outputImage}");
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new Exception("No image thumbnails were extracted.");

        // -------------------------------------------------
        // Optional clean‑up (commented out to keep files for inspection)
        // -------------------------------------------------
        // File.Delete(thumbnailPath);
        // File.Delete(docPath);
    }
}
