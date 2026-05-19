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
        // Prepare working directories.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string inputBmpPath = Path.Combine(workDir, "sample.bmp");
        string docPath = Path.Combine(workDir, "DocumentWithBmp.docx");
        string outputPngDir = Path.Combine(workDir, "Converted");
        Directory.CreateDirectory(outputPngDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic BMP image using Aspose.Drawing.
        // -----------------------------------------------------------------
        int width = 200;
        int height = 200;
        using (Bitmap bmp = new Bitmap(width, height, PixelFormat.Format24bppRgb))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                // Draw a simple gradient rectangle.
                for (int i = 0; i < 10; i++)
                {
                    int shade = 25 * i;
                    Color col = Color.FromArgb(shade, 0, 255 - shade);
                    g.FillRectangle(new SolidBrush(col), i * 20, i * 20, 180 - i * 20, 180 - i * 20);
                }
            }
            bmp.Save(inputBmpPath);
        }

        // -----------------------------------------------------------------
        // 2. Insert the BMP into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputBmpPath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract each BMP image, converting it to PNG
        //    with a reduced color depth of 256 colors (8‑bpp indexed).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original image bytes to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the original BMP using Aspose.Drawing.
                using (Bitmap originalBmp = new Bitmap(originalStream))
                {
                    // Clone the bitmap to an 8‑bpp indexed format.
                    using (Bitmap indexedBmp = originalBmp.Clone(
                        new Rectangle(0, 0, originalBmp.Width, originalBmp.Height),
                        PixelFormat.Format8bppIndexed))
                    {
                        // Save the indexed bitmap as PNG.
                        string outFile = Path.Combine(outputPngDir, $"image_{imageIndex}.png");
                        indexedBmp.Save(outFile, ImageFormat.Png);

                        if (!File.Exists(outFile))
                            throw new InvalidOperationException($"Failed to create output file: {outFile}");
                    }
                }
            }

            imageIndex++;
        }

        // Validate that at least one PNG was produced.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted and converted.");
    }
}
