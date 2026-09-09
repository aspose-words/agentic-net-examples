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
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string outputDir = Path.Combine(artifactsDir, "Output");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image (a simple gradient) using Aspose.Drawing
        // -----------------------------------------------------------------
        string inputImagePath = Path.Combine(artifactsDir, "input.png");
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (var bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        using (var graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            // Fill with a light gray background
            graphics.Clear(Aspose.Drawing.Color.LightGray);
            // Draw a red rectangle
            var redBrush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red);
            graphics.FillRectangle(redBrush, 50, 50, 100, 100);
            // Save the bitmap as PNG
            bitmap.Save(inputImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Insert the sample image into a Word document
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and process each PNG image
        // -----------------------------------------------------------------
        var loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .OfType<Shape>()
                              .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                              .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No PNG images were found in the document.");

        int imageIndex = 0;
        foreach (var shape in shapes)
        {
            // Extract image bytes from the shape
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into an Aspose.Drawing.Bitmap
            using (var ms = new MemoryStream(imageBytes))
            using (var bitmap = new Aspose.Drawing.Bitmap(ms))
            {
                // ---------------------------------------------------------
                // Apply a simple color‑balance adjustment:
                //   - Increase the red channel
                //   - Decrease the blue channel
                // ---------------------------------------------------------
                for (int y = 0; y < bitmap.Height; y++)
                {
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        var pixel = bitmap.GetPixel(x, y);
                        int r = Math.Min(255, pixel.R + 30); // boost red
                        int g = pixel.G;                     // keep green unchanged
                        int b = Math.Max(0, pixel.B - 30);   // reduce blue
                        var newColor = Aspose.Drawing.Color.FromArgb(pixel.A, r, g, b);
                        bitmap.SetPixel(x, y, newColor);
                    }
                }

                // Save the adjusted image to the output folder
                string adjustedPath = Path.Combine(outputDir, $"adjusted_{imageIndex}.png");
                bitmap.Save(adjustedPath);
                imageIndex++;
            }
        }

        // Verify that at least one adjusted image was written
        if (imageIndex == 0)
            throw new InvalidOperationException("No adjusted images were saved.");

        // Optional: clean up the temporary document (not required for the example)
        // File.Delete(docPath);
    }
}
