using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic thumbnail image (sample input for the doc)
        // -----------------------------------------------------------------
        const string thumbnailPath = "thumb.png";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.DarkBlue, 5))
                {
                    g.DrawEllipse(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(thumbnailPath, ImageFormat.Png);
        }

        if (!File.Exists(thumbnailPath))
            throw new InvalidOperationException($"Failed to create thumbnail image at {thumbnailPath}.");

        // -----------------------------------------------------------------
        // 2. Create a new document and insert the thumbnail as an image shape.
        //    (In Aspose.Words there is no dedicated video shape, so we use an
        //     image shape to simulate a video thumbnail.)
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.ImageData.SetImage(thumbnailPath);
        // Optional: give the shape a name that hints it is a video thumbnail.
        imageShape.Name = "VideoThumbnail";

        // Append the shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(imageShape);

        // Save the sample document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save document at {docPath}.");

        // -----------------------------------------------------------------
        // 3. Load the document and extract all image shapes (simulated video thumbnails).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int index = 1;

        foreach (Shape shape in shapeNodes)
        {
            // We look for image shapes that have an image attached.
            if (shape.HasImage)
            {
                string outPath = $"video-thumb-{index}.png";
                shape.ImageData.Save(outPath);
                if (!File.Exists(outPath))
                    throw new InvalidOperationException($"Failed to save extracted image to {outPath}.");

                extractedCount++;
                index++;
            }
        }

        // Validate that at least one thumbnail was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No video thumbnail images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. (Optional) Clean up temporary files.
        // -----------------------------------------------------------------
        // File.Delete(thumbnailPath);
        // File.Delete(docPath);
    }
}
