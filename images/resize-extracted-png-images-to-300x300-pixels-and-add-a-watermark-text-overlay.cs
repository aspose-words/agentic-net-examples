using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeDrawing = Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a deterministic folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image (500x500) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string inputImagePath = Path.Combine(artifactsDir, "input.png");
        using (AsposeDrawing.Bitmap bitmap = new AsposeDrawing.Bitmap(500, 500))
        {
            using (AsposeDrawing.Graphics graphics = AsposeDrawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(AsposeDrawing.Color.White);
                // Draw a simple ellipse for visual content.
                graphics.DrawEllipse(AsposeDrawing.Pens.Black, 50, 50, 400, 400);
            }
            bitmap.Save(inputImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Insert the sample image into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract PNG images from the document.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the extracted PNG to a temporary file.
            string extractedPath = Path.Combine(artifactsDir, $"extracted_{extractedCount}.png");
            shape.ImageData.Save(extractedPath);

            // -----------------------------------------------------------------
            // 4. Resize the extracted image to 300x300 and add a watermark.
            // -----------------------------------------------------------------
            using (AsposeDrawing.Bitmap original = new AsposeDrawing.Bitmap(extractedPath))
            {
                using (AsposeDrawing.Bitmap resized = new AsposeDrawing.Bitmap(300, 300))
                {
                    using (AsposeDrawing.Graphics graphics = AsposeDrawing.Graphics.FromImage(resized))
                    {
                        // Ensure a clean canvas.
                        graphics.Clear(AsposeDrawing.Color.Transparent);

                        // Draw the original image scaled to 300x300.
                        graphics.DrawImage(
                            original,
                            new AsposeDrawing.Rectangle(0, 0, 300, 300));

                        // Prepare watermark text.
                        using (AsposeDrawing.Font watermarkFont = new AsposeDrawing.Font("Arial", 24))
                        {
                            // Semi‑transparent white brush.
                            using (AsposeDrawing.SolidBrush brush = new AsposeDrawing.SolidBrush(
                                AsposeDrawing.Color.FromArgb(128, AsposeDrawing.Color.White)))
                            {
                                // Position the watermark near the bottom‑left corner.
                                graphics.DrawString(
                                    "Watermark",
                                    watermarkFont,
                                    brush,
                                    new AsposeDrawing.PointF(10, 260));
                            }
                        }
                    }

                    // Save the watermarked image.
                    string watermarkedPath = Path.Combine(artifactsDir, $"watermarked_{extractedCount}.png");
                    resized.Save(watermarkedPath);

                    // Validate that the output file exists.
                    if (!File.Exists(watermarkedPath))
                        throw new InvalidOperationException("Watermarked image was not created.");
                }
            }

            extractedCount++;
        }

        // Ensure at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were found in the document.");

        // The example finishes without requiring user interaction.
    }
}
