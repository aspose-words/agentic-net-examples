using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        // Fill background with a light gray color.
        graphics.Clear(Aspose.Drawing.Color.LightGray);
        // Draw a simple red ellipse to have visible content.
        using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
        {
            graphics.DrawEllipse(pen, 20, 20, 160, 160);
        }
        // Save the bitmap as PNG.
        bitmap.Save(sampleImagePath);
        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the sample PNG image twice.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.Writeln(); // Add a line break between images.
        builder.InsertImage(sampleImagePath);

        // -----------------------------------------------------------------
        // 3. Extract all PNG images, enhance contrast, and save them.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Ensure the shape actually contains an image.
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Apply maximum contrast (value range 0.0 to 1.0).
            shape.ImageData.Contrast = 1.0;

            // Save the enhanced image to a deterministic file name.
            string outFile = Path.Combine(artifactsDir, $"extracted_{extractedCount}.png");
            shape.ImageData.Save(outFile);

            // Verify that the file was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Failed to save image '{outFile}'.");

            extractedCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validate that at least one PNG image was processed.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        // Optional: Save the document (not required for the task but demonstrates full workflow).
        string docPath = Path.Combine(artifactsDir, "DocumentWithImages.docx");
        doc.Save(docPath);
    }
}
