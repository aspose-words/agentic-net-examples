using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Create a new Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string originalDocPath = "original.docx";
        doc.Save(originalDocPath);

        // Load the document (simulating a separate load step).
        Document loadedDoc = new Document(originalDocPath);

        // Process all image shapes, apply motion blur to JPEG images, and replace them.
        bool anyImageProcessed = false;
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Apply a simple horizontal motion blur.
                    using (Bitmap blurredBitmap = ApplyMotionBlur(originalBitmap))
                    {
                        // Save the blurred bitmap to a temporary JPEG file.
                        string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".jpg");
                        blurredBitmap.Save(tempPath, ImageFormat.Jpeg);

                        // Replace the shape's image with the blurred version from file.
                        shape.ImageData.SetImage(tempPath);
                        anyImageProcessed = true;

                        // Clean up the temporary file.
                        if (File.Exists(tempPath))
                            File.Delete(tempPath);
                    }
                }
            }
        }

        if (!anyImageProcessed)
            throw new Exception("No JPEG images were found to process.");

        // Save the modified document.
        const string blurredDocPath = "blurred.docx";
        loadedDoc.Save(blurredDocPath);

        // Validate that the output file exists.
        if (!File.Exists(blurredDocPath))
            throw new Exception("Failed to create the output document.");

        // Optional cleanup.
        // File.Delete(sampleImagePath);
        // File.Delete(originalDocPath);
    }

    // Creates a simple deterministic JPEG image.
    private static void CreateSampleJpeg(string path)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 5))
                {
                    g.DrawEllipse(pen, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    // Applies a basic horizontal motion blur by drawing the image multiple times with offsets.
    private static Bitmap ApplyMotionBlur(Bitmap source)
    {
        const int blurLength = 10; // Number of pixels to blur.
        Bitmap dest = new Bitmap(source.Width, source.Height);
        using (Graphics g = Graphics.FromImage(dest))
        {
            // Clear destination.
            g.Clear(Aspose.Drawing.Color.Transparent);

            // Draw the source image multiple times with incremental horizontal offsets.
            for (int offset = -blurLength; offset <= blurLength; offset++)
            {
                g.DrawImage(source, offset, 0);
            }
        }
        return dest;
    }
}
