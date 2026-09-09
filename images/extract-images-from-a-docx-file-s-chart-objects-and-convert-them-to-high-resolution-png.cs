using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractChartImages
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare output folder.
        // -----------------------------------------------------------------
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 2. Create a deterministic sample image (white background with a black rectangle).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Black, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 3. Create a DOCX and insert the sample image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample document with an image (standing in for a chart):");
        builder.InsertImage(sampleImagePath);

        string sourceDocPath = Path.Combine(artifactsDir, "ImageDocument.docx");
        doc.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 4. Load the document and extract images from shape objects.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>();

        int extractedCount = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                // Determine file extension based on the original image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(artifactsDir, $"ExtractedImage_{extractedCount}{extension}");

                // Save the original image data (preserves original format).
                shape.ImageData.Save(outFile);
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from shape objects.");

        // -----------------------------------------------------------------
        // 5. Render each shape containing an image to a high‑resolution PNG.
        // -----------------------------------------------------------------
        int renderedCount = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string renderPath = Path.Combine(artifactsDir, $"RenderedShape_{renderedCount}.png");

                ImageSaveOptions renderOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    // Set both horizontal and vertical DPI to 300.
                    Resolution = 300
                };

                // Render the shape to a PNG file with the specified resolution.
                shape.GetShapeRenderer().Save(renderPath, renderOptions);
                renderedCount++;
            }
        }

        // -----------------------------------------------------------------
        // 6. Completion message.
        // -----------------------------------------------------------------
        Console.WriteLine($"Extracted {extractedCount} image(s) and rendered {renderedCount} shape(s) to '{artifactsDir}'.");
    }
}
