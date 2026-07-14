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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a deterministic PNG image to be used in the document.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(sampleImagePath, 200, 200);

        // Build a document that contains the PNG image twice.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph();
        builder.InsertImage(sampleImagePath);

        // (Optional) Save the document for reference.
        string docPath = Path.Combine(artifactsDir, "DocumentWithImages.docx");
        doc.Save(docPath);

        // Extract PNG images, enhance contrast, and save them.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        int savedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Apply maximum contrast (range 0.0 to 1.0).
            shape.ImageData.Contrast = 1.0;

            // Save the enhanced image.
            string outputImagePath = Path.Combine(artifactsDir, $"extracted_{imageIndex}.png");
            shape.ImageData.Save(outputImagePath);
            savedCount++;
            imageIndex++;
        }

        // Validate that at least one image was saved.
        if (savedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and saved.");
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        // Create a bitmap.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Obtain a graphics object for drawing.
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);

                // Draw a solid red rectangle.
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
                {
                    graphics.FillRectangle(brush, 20, 20, width - 40, height - 40);
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath);
        }
    }
}
