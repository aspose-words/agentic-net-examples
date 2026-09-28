using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Drawing2D;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample TIFF image using Aspose.Drawing.
        // ------------------------------------------------------------
        const string tiffPath = "sample.tiff";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            // Save the bitmap as a TIFF file.
            bitmap.Save(tiffPath, ImageFormat.Tiff);
        }

        // ------------------------------------------------------------
        // 2. Insert the TIFF image into a Word document.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(tiffPath);
        const string docPath = "DocumentWithTiff.docx";
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Extract images, convert each to grayscale JPEG.
        // ------------------------------------------------------------
        List<string> outputFiles = new List<string>();
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue; // Skip non‑image shapes.

            // Load the image data into a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset for reading.

                // Load the source bitmap.
                using (Bitmap sourceBitmap = new Bitmap(imageStream))
                {
                    // Create a new bitmap for the grayscale result.
                    using (Bitmap grayBitmap = new Bitmap(sourceBitmap.Width, sourceBitmap.Height))
                    {
                        using (Graphics graphics = Graphics.FromImage(grayBitmap))
                        {
                            // Process each pixel to compute its grayscale value.
                            for (int y = 0; y < sourceBitmap.Height; y++)
                            {
                                for (int x = 0; x < sourceBitmap.Width; x++)
                                {
                                    Color pixel = sourceBitmap.GetPixel(x, y);
                                    int gray = (int)(pixel.R * 0.3 + pixel.G * 0.59 + pixel.B * 0.11);
                                    Color grayColor = Color.FromArgb(gray, gray, gray);
                                    grayBitmap.SetPixel(x, y, grayColor);
                                }
                            }
                        }

                        // Save the grayscale bitmap as JPEG.
                        string outputPath = $"image-{outputFiles.Count + 1}.jpg";
                        grayBitmap.Save(outputPath, ImageFormat.Jpeg);
                        outputFiles.Add(outputPath);
                    }
                }
            }
        }

        // ------------------------------------------------------------
        // 4. Validate that at least one JPEG was created.
        // ------------------------------------------------------------
        if (outputFiles.Count == 0)
            throw new InvalidOperationException("No images were found or converted.");

        // Optional cleanup (uncomment if desired).
        // File.Delete(tiffPath);
        // File.Delete(docPath);
    }
}
