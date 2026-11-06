using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic BMP image using Aspose.Drawing.
        const string bmpPath = "sample.bmp";
        Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bmp);
        g.Clear(Aspose.Drawing.Color.White);
        // Draw a simple red rectangle.
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
        {
            g.FillRectangle(brush, new Aspose.Drawing.Rectangle(50, 50, 100, 100));
        }
        bmp.Save(bmpPath, Aspose.Drawing.Imaging.ImageFormat.Bmp);
        g.Dispose();
        bmp.Dispose();

        // Step 2: Insert the BMP image into a Word document.
        const string docPath = "DocumentWithBmp.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        doc.Save(docPath);

        // Step 3: Load the document and convert each extracted image to PNG with 256‑color depth.
        Document loadedDoc = new Document(docPath);
        List<string> convertedFiles = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage)
                continue;

            // Save the image data to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Aspose.Drawing.Bitmap originalBmp = new Aspose.Drawing.Bitmap(imgStream))
                {
                    // Convert to 8‑bpp indexed (256 colors).
                    using (Aspose.Drawing.Bitmap indexedBmp = originalBmp.Clone(
                        new Aspose.Drawing.Rectangle(0, 0, originalBmp.Width, originalBmp.Height),
                        Aspose.Drawing.Imaging.PixelFormat.Format8bppIndexed))
                    {
                        // Save as PNG.
                        string pngPath = $"converted-{imageIndex}.png";
                        indexedBmp.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);
                        convertedFiles.Add(pngPath);
                        imageIndex++;
                    }
                }
            }
        }

        // Step 4: Validate that at least one PNG file was created.
        if (convertedFiles.Count == 0)
        {
            throw new InvalidOperationException("No images were found or converted.");
        }

        // Optional cleanup (comment out to keep files for inspection).
        // File.Delete(bmpPath);
        // File.Delete(docPath);
    }
}
