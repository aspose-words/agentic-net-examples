using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample GIF image (single‑frame for simplicity).
        // -----------------------------------------------------------------
        string sampleGifPath = Path.Combine(artifactsDir, "sample.gif");
        const int width = 200;
        const int height = 200;

        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple rectangle.
        graphics.FillRectangle(new SolidBrush(Color.Blue), 20, 20, width - 40, height - 40);
        // Save as GIF.
        bitmap.Save(sampleGifPath, ImageFormat.Gif);
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Insert the GIF into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleGifPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithGif.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract all GIF images from the document and convert them to PNG.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int convertedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Gif)
                continue;

            // Save the original GIF (optional, just to demonstrate extraction).
            string extractedGifPath = Path.Combine(artifactsDir, $"extracted_{convertedCount}.gif");
            shape.ImageData.Save(extractedGifPath);

            // Load the GIF into Aspose.Drawing.Image.
            using (MemoryStream gifStream = new MemoryStream())
            {
                shape.ImageData.Save(gifStream);
                gifStream.Position = 0;

                using (Aspose.Drawing.Image gifImage = Aspose.Drawing.Image.FromStream(gifStream))
                {
                    // Convert to PNG. (Aspose.Drawing saves only the first frame for GIFs;
                    // for a true animated PNG you would need a library that supports APNG.)
                    string pngPath = Path.Combine(artifactsDir, $"converted_{convertedCount}.png");
                    gifImage.Save(pngPath, ImageFormat.Png);
                }
            }

            convertedCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one PNG was produced.
        // -----------------------------------------------------------------
        if (convertedCount == 0)
            throw new InvalidOperationException("No GIF images were found to convert.");

        // Verify that the PNG files exist.
        for (int i = 0; i < convertedCount; i++)
        {
            string pngPath = Path.Combine(artifactsDir, $"converted_{i}.png");
            if (!File.Exists(pngPath))
                throw new FileNotFoundException($"Expected output file not found: {pngPath}");
        }

        // The example finishes without requiring user interaction.
    }
}
