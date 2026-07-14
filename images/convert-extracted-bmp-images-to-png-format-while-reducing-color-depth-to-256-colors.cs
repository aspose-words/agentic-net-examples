using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic BMP image using Aspose.Drawing.
        string bmpPath = Path.Combine(artifactsDir, "sample.bmp");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Brush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(255, 0, 120, 215)))
                {
                    g.FillRectangle(brush, 20, 20, 160, 160);
                }
            }
            bmp.Save(bmpPath, ImageFormat.Bmp);
        }

        // 2. Insert the BMP into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithBmp.docx");
        doc.Save(docPath);

        // 3. Load the document and extract each image.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Get raw image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap originalBmp = new Bitmap(ms))
                {
                    // Draw the original bitmap onto a non‑indexed bitmap first.
                    using (Bitmap tempBmp = new Bitmap(originalBmp.Width, originalBmp.Height, PixelFormat.Format24bppRgb))
                    {
                        using (Graphics g = Graphics.FromImage(tempBmp))
                        {
                            g.DrawImage(originalBmp, 0, 0, originalBmp.Width, originalBmp.Height);
                        }

                        // Clone the temporary bitmap to an 8‑bpp indexed bitmap (256 colors).
                        using (Bitmap indexedBmp = tempBmp.Clone(
                            new Rectangle(0, 0, tempBmp.Width, tempBmp.Height),
                            PixelFormat.Format8bppIndexed))
                        {
                            // Save the indexed bitmap as PNG.
                            string pngPath = Path.Combine(artifactsDir,
                                $"ExtractedImage_{imageIndex}.png");
                            indexedBmp.Save(pngPath, ImageFormat.Png);

                            // Validate that the PNG file was created.
                            if (!File.Exists(pngPath))
                                throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");
                        }
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found to convert.");
    }
}
