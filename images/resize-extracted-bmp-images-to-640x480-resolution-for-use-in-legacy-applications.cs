using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string sampleBmpPath = "sample.bmp";
        const string docPath = "DocumentWithBmp.docx";

        // -------------------------------------------------
        // 1. Create a sample BMP image (200x200, solid blue)
        // -------------------------------------------------
        Aspose.Drawing.Bitmap sampleBitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics sampleGraphics = Aspose.Drawing.Graphics.FromImage(sampleBitmap);
        sampleGraphics.Clear(Aspose.Drawing.Color.Blue);
        sampleGraphics.Dispose();
        sampleBitmap.Save(sampleBmpPath);
        sampleBitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a Word document and insert the BMP image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image from the file path (preferred workflow)
        builder.InsertImage(sampleBmpPath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images, resizing each to 640x480
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Get the raw image bytes from the shape
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure stream is at the beginning
                // Load the original image using Aspose.Drawing
                using (Aspose.Drawing.Bitmap originalBmp = new Aspose.Drawing.Bitmap(ms))
                {
                    // Create a new bitmap with the target size 640x480
                    using (Aspose.Drawing.Bitmap resizedBmp = new Aspose.Drawing.Bitmap(640, 480))
                    {
                        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(resizedBmp))
                        {
                            // Fill background (optional)
                            g.Clear(Aspose.Drawing.Color.White);
                            // Draw the original image scaled to the new size
                            g.DrawImage(
                                originalBmp,
                                new Aspose.Drawing.Rectangle(0, 0, 640, 480));
                        }

                        // Save the resized BMP to a deterministic file name
                        string resizedPath = $"resized_{extractedCount}.bmp";
                        resizedBmp.Save(resizedPath);
                        // Validate that the file was created
                        if (!File.Exists(resizedPath))
                            throw new InvalidOperationException($"Failed to save resized image '{resizedPath}'.");
                        extractedCount++;
                    }
                }
            }
        }

        // -------------------------------------------------
        // 4. Validation: ensure at least one image was resized
        // -------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted and resized.");

        // Optional cleanup (commented out for debugging purposes)
        // File.Delete(sampleBmpPath);
        // File.Delete(docPath);
    }
}
