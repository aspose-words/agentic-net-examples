using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputImagePath = Path.Combine(artifactsDir, "input.png");

        // -------------------------------------------------
        // 1. Create a deterministic sample PNG image.
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill background with white.
            g.Clear(Aspose.Drawing.Color.White);
            // Draw a red rectangle.
            using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                g.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
            }
            // Save the bitmap as PNG.
            bitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the PNG image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);

        // -------------------------------------------------
        // 3. Extract all PNG images, apply a simple color‑balance
        //    adjustment, and save the adjusted images.
        // -------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Only process PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Get the raw image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure the stream is at the beginning.
                using (Bitmap bmp = new Bitmap(ms))
                {
                    // Simple color‑balance: increase Red, decrease Blue.
                    for (int y = 0; y < bmp.Height; y++)
                    {
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            Aspose.Drawing.Color orig = bmp.GetPixel(x, y);
                            int r = Math.Min(255, orig.R + 30);
                            int g = orig.G;
                            int b = Math.Max(0, orig.B - 30);
                            Aspose.Drawing.Color adjusted = Aspose.Drawing.Color.FromArgb(orig.A, r, g, b);
                            bmp.SetPixel(x, y, adjusted);
                        }
                    }

                    // Determine output file name with proper extension.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outputPath = Path.Combine(artifactsDir, $"adjusted_{imageIndex}{extension}");

                    // Save the adjusted bitmap.
                    bmp.Save(outputPath);
                    imageIndex++;
                }
            }
        }

        // -------------------------------------------------
        // 4. Validation – ensure at least one image was saved.
        // -------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted and processed.");

        // Optional: list saved files (not required for execution).
        // foreach (var file in Directory.GetFiles(artifactsDir, "adjusted_*.png"))
        //     Console.WriteLine($"Saved: {file}");
    }
}
