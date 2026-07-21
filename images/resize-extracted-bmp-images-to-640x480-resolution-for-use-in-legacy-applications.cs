using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing namespace
using Aspose.Drawing.Imaging;      // For ImageFormat

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a deterministic sample BMP image (200x200, solid blue).
        // ------------------------------------------------------------
        const string sampleBmpPath = "sample.bmp";

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.Blue);
            // Save as BMP – the file extension determines the format.
            bitmap.Save(sampleBmpPath, ImageFormat.Bmp);
        }

        // ------------------------------------------------------------
        // 2. Create a new Word document and insert the sample BMP image.
        // ------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        const string docPath = "docWithBmp.docx";
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Reload the document (demonstrates the load step).
        // ------------------------------------------------------------
        var loadedDoc = new Document(docPath);

        // ------------------------------------------------------------
        // 4. Extract images, resize them to 640x480, and save the resized versions.
        //    The original example filtered only BMP images; Aspose.Words may
        //    convert BMP to another format on insert, so we process any image.
        // ------------------------------------------------------------
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int resizedCount = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream.
            using (var originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the original image using Aspose.Drawing.
                using (var originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    // Create a new bitmap with the target size.
                    using (var resizedBitmap = new Aspose.Drawing.Bitmap(640, 480))
                    using (var g = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                    {
                        // Optional: fill background with white.
                        g.Clear(Aspose.Drawing.Color.White);
                        // Draw the original image stretched to the new dimensions.
                        g.DrawImage(originalBitmap, new Aspose.Drawing.Rectangle(0, 0, 640, 480));

                        // Save the resized BMP to a deterministic file name.
                        string resizedPath = $"resized_{resizedCount}.bmp";
                        resizedBitmap.Save(resizedPath, ImageFormat.Bmp);
                        resizedCount++;
                    }
                }
            }
        }

        // ------------------------------------------------------------
        // 5. Validate that at least one image was resized.
        // ------------------------------------------------------------
        if (resizedCount == 0)
            throw new InvalidOperationException("No images were found to resize.");
    }
}
