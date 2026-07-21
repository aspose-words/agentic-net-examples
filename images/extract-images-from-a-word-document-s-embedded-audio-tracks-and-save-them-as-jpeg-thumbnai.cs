using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractImageThumbnails
{
    public static void Main()
    {
        // Folder for all generated files
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image that will be inserted into the document.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        int imgWidth = 200;
        int imgHeight = 200;

        // Use Aspose.Drawing to create a deterministic sample image.
        using (Bitmap bmp = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            // Fill background with white and draw a simple rectangle.
            g.Clear(Color.White);
            using (Pen pen = new Pen(Color.Blue, 5))
            {
                g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
            }
            // Save the image to disk.
            bmp.Save(sampleImagePath);
        }

        // ---------------------------------------------------------------
        // 2. Create a Word document and insert the sample image into it.
        // ---------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(outputDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 3. Load the document and extract each image, creating a JPEG thumbnail.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a temporary file.
            string originalImagePath = Path.Combine(outputDir,
                $"Image_{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
            shape.ImageData.Save(originalImagePath);

            // Load the saved image using Aspose.Drawing.
            using (Bitmap originalBitmap = new Bitmap(originalImagePath))
            {
                // Define thumbnail size.
                int thumbWidth = 100;
                int thumbHeight = 100;

                using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                using (Graphics g = Graphics.FromImage(thumbBitmap))
                {
                    // Draw the scaled image onto the thumbnail bitmap.
                    g.Clear(Color.Transparent);
                    g.DrawImage(originalBitmap, new Rectangle(0, 0, thumbWidth, thumbHeight));

                    // Save the thumbnail as JPEG.
                    string thumbPath = Path.Combine(outputDir,
                        $"Thumbnail_{imageIndex}.jpg");
                    thumbBitmap.Save(thumbPath, ImageFormat.Jpeg);
                }
            }

            imageIndex++;
        }

        // Simple validation: ensure at least one thumbnail was created.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found in the document.");

        Console.WriteLine($"Extraction complete. Thumbnails saved to: {outputDir}");
    }
}
