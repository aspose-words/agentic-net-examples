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
        // Prepare output folder
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample BMP image using Aspose.Drawing
        string bmpPath = Path.Combine(artifactsDir, "sample.bmp");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Fill with a vertical gradient to make color reduction visible
                for (int y = 0; y < 200; y++)
                {
                    int red = (y * 255) / 199;
                    using (SolidBrush brush = new SolidBrush(Color.FromArgb(red, 0, 255 - red)))
                    {
                        g.FillRectangle(brush, 0, y, 200, 1);
                    }
                }
            }
            bmp.Save(bmpPath, ImageFormat.Bmp);
        }

        // 2. Insert the BMP image into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        doc.Save(docPath);

        // 3. Load the document and extract images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue; // Skip shapes without image data

            // Save the original image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing
                using (Bitmap originalBmp = new Bitmap(originalStream))
                {
                    // Clone the bitmap to an 8‑bpp indexed format (256 colors)
                    using (Bitmap indexedBmp = originalBmp.Clone(
                        new Rectangle(0, 0, originalBmp.Width, originalBmp.Height),
                        PixelFormat.Format8bppIndexed))
                    {
                        // Save the indexed bitmap as PNG
                        string pngPath = Path.Combine(artifactsDir, $"converted_{imageIndex}.png");
                        indexedBmp.Save(pngPath, ImageFormat.Png);

                        // Validate that the PNG file was created
                        if (!File.Exists(pngPath))
                            throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found and converted.");
    }
}
