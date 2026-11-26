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
        // ------------------------------------------------------------
        // 1. Create a deterministic BMP image using Aspose.Drawing.
        // ------------------------------------------------------------
        const string bmpPath = "sample.bmp";
        Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bmp);
        graphics.Clear(Aspose.Drawing.Color.White);
        graphics.DrawRectangle(Aspose.Drawing.Pens.Black, 20, 20, 160, 160);
        // Save the BMP file – this file will be inserted into the Word document.
        bmp.Save(bmpPath, Aspose.Drawing.Imaging.ImageFormat.Bmp);
        graphics.Dispose();
        bmp.Dispose();

        // ------------------------------------------------------------
        // 2. Insert the BMP image into a new Word document.
        // ------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Reload the document and enumerate all Shape nodes that contain images.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int convertedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue; // Skip shapes without images.

            // ------------------------------------------------------------
            // 4. Extract the image data to a memory stream.
            // ------------------------------------------------------------
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset for reading.

                // ------------------------------------------------------------
                // 5. Load the extracted image into Aspose.Drawing.Bitmap.
                // ------------------------------------------------------------
                using (Aspose.Drawing.Bitmap originalBmp = new Aspose.Drawing.Bitmap(imageStream))
                {
                    // ------------------------------------------------------------
                    // 6. Reduce color depth to 256 colors (8‑bpp indexed).
                    // ------------------------------------------------------------
                    using (Aspose.Drawing.Bitmap indexedBmp = originalBmp.Clone(
                        new Rectangle(0, 0, originalBmp.Width, originalBmp.Height),
                        Aspose.Drawing.Imaging.PixelFormat.Format8bppIndexed))
                    {
                        // ------------------------------------------------------------
                        // 7. Save the converted image as PNG.
                        // ------------------------------------------------------------
                        string pngPath = $"converted-{convertedCount}.png";
                        indexedBmp.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);

                        // Validate that the PNG file was created.
                        if (!File.Exists(pngPath))
                            throw new Exception($"Failed to create PNG file: {pngPath}");

                        convertedCount++;
                    }
                }
            }
        }

        // ------------------------------------------------------------
        // 8. Ensure at least one image was converted.
        // ------------------------------------------------------------
        if (convertedCount == 0)
            throw new Exception("No images were found and converted.");

        // Optional cleanup (commented out to allow inspection of generated files).
        // File.Delete(bmpPath);
        // File.Delete(docPath);
    }
}
