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
        // Prepare input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic BMP image using Aspose.Drawing.
        string bmpPath = Path.Combine(inputDir, "sample.bmp");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.White);
                g.FillEllipse(new SolidBrush(Aspose.Drawing.Color.Blue), 20, 20, 160, 160);
            }
            // Save as BMP.
            bmp.Save(bmpPath, ImageFormat.Bmp);
        }

        // 2. Insert the BMP image into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        string docPath = Path.Combine(inputDir, "sample.docx");
        doc.Save(docPath);

        // 3. Load the document and extract images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue; // Skip shapes without images.

            // Retrieve raw image bytes from the shape.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image bytes into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0;
                using (Bitmap originalBmp = new Bitmap(ms))
                {
                    // Clone the bitmap to an 8‑bpp indexed format (256 colors).
                    using (Bitmap indexedBmp = originalBmp.Clone(
                        new Rectangle(0, 0, originalBmp.Width, originalBmp.Height),
                        PixelFormat.Format8bppIndexed))
                    {
                        // Save the converted image as PNG.
                        string pngPath = Path.Combine(outputDir, $"image_{imageIndex}.png");
                        indexedBmp.Save(pngPath, ImageFormat.Png);

                        // Validate that the PNG file was created.
                        if (!File.Exists(pngPath))
                            throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");
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
