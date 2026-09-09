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
        // Prepare folders.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleImagePath, 200, 200);

        // 2. Create a Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string originalDocPath = Path.Combine(artifactsDir, "original.docx");
        doc.Save(originalDocPath);

        // 3. Load the document, extract JPEG images, apply motion blur, and re‑embed them.
        Document loadedDoc = new Document(originalDocPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                              .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                              .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        foreach (var shape in shapes)
        {
            // Extract the image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Apply a simple horizontal motion blur.
                    using (Bitmap blurredBitmap = ApplyHorizontalMotionBlur(originalBitmap))
                    {
                        // Save the blurred bitmap back to a stream.
                        using (MemoryStream blurredStream = new MemoryStream())
                        {
                            blurredBitmap.Save(blurredStream, ImageFormat.Jpeg);
                            blurredStream.Position = 0;

                            // Replace the shape's image with the blurred version.
                            shape.ImageData.SetImage(blurredStream);
                        }
                    }
                }
            }
        }

        // 4. Save the modified document.
        string processedDocPath = Path.Combine(artifactsDir, "processed.docx");
        loadedDoc.Save(processedDocPath);

        // Validation.
        if (!File.Exists(processedDocPath))
            throw new InvalidOperationException("The processed document was not saved.");

        Console.WriteLine("Processing complete. Files are located in: " + artifactsDir);
    }

    // Creates a deterministic JPEG image with simple graphics.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a red ellipse.
            graphics.FillEllipse(new SolidBrush(Color.Red), 20, 20, width - 40, height - 40);
            // Save as JPEG.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Applies a basic horizontal motion blur by averaging neighboring pixels.
    private static Bitmap ApplyHorizontalMotionBlur(Bitmap source)
    {
        int w = source.Width;
        int h = source.Height;
        Bitmap result = new Bitmap(w, h);

        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                // Gather colors of the current pixel and its immediate horizontal neighbours.
                Color cCenter = source.GetPixel(x, y);
                Color cLeft = x > 0 ? source.GetPixel(x - 1, y) : cCenter;
                Color cRight = x < w - 1 ? source.GetPixel(x + 1, y) : cCenter;

                // Average the RGB components.
                int r = (cLeft.R + cCenter.R + cRight.R) / 3;
                int g = (cLeft.G + cCenter.G + cRight.G) / 3;
                int b = (cLeft.B + cCenter.B + cRight.B) / 3;

                result.SetPixel(x, y, Color.FromArgb(r, g, b));
            }
        }

        return result;
    }
}
