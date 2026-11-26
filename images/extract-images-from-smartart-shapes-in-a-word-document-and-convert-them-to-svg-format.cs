using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic raster image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        try
        {
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                // Fill background with white.
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a simple red rectangle.
                graphics.FillRectangle(new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red), 20, 20, 160, 160);
            }
            finally
            {
                graphics.Dispose();
            }

            // Save the bitmap to a PNG file.
            bitmap.Save(sampleImagePath);
        }
        finally
        {
            bitmap.Dispose();
        }

        // -----------------------------------------------------------------
        // 2. Create a new Word document and insert the sample image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as a shape (inline by default).
        Shape imageShape = builder.InsertImage(sampleImagePath);
        // Optionally, set a size for the shape.
        imageShape.Width = 200;
        imageShape.Height = 200;

        // Save the document (optional, just to have a source file).
        string docPath = Path.Combine(outputDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Render each shape (including SmartArt) to SVG.
        // -----------------------------------------------------------------
        int svgIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Ensure SmartArt drawings are up‑to‑date.
            shape.UpdateSmartArtDrawing();

            // Prepare SVG save options – specify a resources folder to avoid
            // the “Resource file(s) cannot be written to disk” exception.
            SvgSaveOptions svgOptions = new SvgSaveOptions
            {
                ExportEmbeddedImages = false,
                ResourcesFolder = outputDir
            };

            string svgPath = Path.Combine(outputDir, $"ExtractedShape_{svgIndex}.svg");
            ShapeRenderer renderer = shape.GetShapeRenderer();
            renderer.Save(svgPath, svgOptions);
            svgIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Validate that at least one SVG file was created.
        // -----------------------------------------------------------------
        if (svgIndex == 0 || !File.Exists(Path.Combine(outputDir, "ExtractedShape_0.svg")))
            throw new InvalidOperationException("No SVG files were generated.");
    }
}
