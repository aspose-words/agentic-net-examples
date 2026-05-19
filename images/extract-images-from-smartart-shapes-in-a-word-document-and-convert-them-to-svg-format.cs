using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // For deterministic bitmap creation

public class ExtractSmartArtToSvg
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare output folder.
        // -----------------------------------------------------------------
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 2. Create a deterministic sample PNG image.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // -----------------------------------------------------------------
        // 3. Build a document that contains the image (simulating a SmartArt shape).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape imageShape = builder.InsertImage(sampleImagePath);
        imageShape.Width = 400;
        imageShape.Height = 400;

        // Save the document (optional, for reference).
        string docPath = Path.Combine(outputDir, "SampleDocument.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 4. Load the document and render each shape that contains an image to SVG.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int svgIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // For SmartArt shapes, ensure the pre‑rendered drawing is up‑to‑date.
            // This call is safe for non‑SmartArt shapes as well.
            shape.UpdateSmartArtDrawing();

            // Only shapes that actually have image data can be rendered.
            if (shape.HasImage)
            {
                // Configure SVG save options.
                SvgSaveOptions svgOptions = new SvgSaveOptions
                {
                    ExportEmbeddedImages = false,
                    // When ExportEmbeddedImages is false Aspose.Words needs a folder to write the
                    // raster resources (e.g., PNGs) that belong to the SVG.
                    ResourcesFolder = outputDir,
                    ResourcesFolderAlias = outputDir
                };

                string svgFileName = Path.Combine(outputDir, $"Shape_{svgIndex}.svg");
                shape.GetShapeRenderer().Save(svgFileName, svgOptions);
                svgIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 5. Validate that at least one SVG file was created.
        // -----------------------------------------------------------------
        if (svgIndex == 0)
            throw new InvalidOperationException("No image shapes were found or SVG files were not created.");
    }

    // Helper method to create a deterministic sample PNG image.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Use Aspose.Drawing types as required by the rule set.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        try
        {
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple blue rectangle border.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
            }
            finally
            {
                graphics.Dispose();
            }

            // Save the bitmap as PNG.
            bitmap.Save(filePath);
        }
        finally
        {
            bitmap.Dispose();
        }
    }
}
