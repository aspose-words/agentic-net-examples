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
        // Step 1: Create a deterministic PNG image file.
        const string pngPath = "sample.png";
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual content.
                g.FillRectangle(new SolidBrush(Color.Blue), 20, 20, imgWidth - 40, imgHeight - 40);
            }
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Step 2: Create a Word document and insert the PNG image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        doc.Save(docPath);

        // Step 3: Load the document and extract PNG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the image to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset stream position before reading.

                // Load the PNG into a bitmap.
                using (Bitmap pngBitmap = new Bitmap(ms))
                {
                    // Preserve original dimensions when saving as JPEG.
                    string jpegPath = $"extracted-{imageIndex}.jpg";
                    pngBitmap.Save(jpegPath, ImageFormat.Jpeg);

                    // Validate that the JPEG file was created.
                    if (!File.Exists(jpegPath))
                        throw new InvalidOperationException($"Failed to create JPEG file: {jpegPath}");

                    imageIndex++;
                }
            }
        }

        // Final validation: ensure at least one JPEG was produced.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found and converted.");

        // Cleanup: optional removal of intermediate files (commented out).
        // File.Delete(pngPath);
        // File.Delete(docPath);
    }
}
