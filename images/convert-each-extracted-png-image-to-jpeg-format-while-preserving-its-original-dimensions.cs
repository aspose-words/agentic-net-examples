using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic PNG sample image.
        const string pngPath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.LightBlue);
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Build a document that contains the PNG image twice.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        builder.Writeln(); // separate the images
        builder.InsertImage(pngPath);
        const string docPath = "DocWithImages.docx";
        doc.Save(docPath);

        // Reload the document (optional, demonstrates load usage).
        Document loadedDoc = new Document(docPath);

        // Extract each PNG image, convert to JPEG while preserving dimensions.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int jpegIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Png) continue;

            // Save the PNG image to a memory stream.
            using (MemoryStream pngStream = new MemoryStream())
            {
                shape.ImageData.Save(pngStream);
                pngStream.Position = 0;

                // Load the PNG into a bitmap and save as JPEG.
                using (Bitmap bitmap = new Bitmap(pngStream))
                {
                    string jpegPath = $"extracted_{jpegIndex}.jpg";
                    bitmap.Save(jpegPath, ImageFormat.Jpeg);
                    jpegIndex++;
                }
            }
        }

        // Validation: ensure at least one JPEG was created.
        if (jpegIndex == 0)
            throw new InvalidOperationException("No PNG images were found to convert.");

        // Cleanup sample files (optional).
        // File.Delete(pngPath);
        // File.Delete(docPath);
    }
}
