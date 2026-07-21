using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;

public class ExtractChartImages
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX containing a chart.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart.
        builder.InsertChart(ChartType.Column, 400, 300);
        doc.Save(Path.Combine(artifactsDir, "ChartDocument.docx"));

        // -----------------------------------------------------------------
        // 2. Load the document and locate chart shapes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(Path.Combine(artifactsDir, "ChartDocument.docx"));
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int chartCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Chart objects are represented by shapes that contain a Chart.
            if (shape.Chart != null)
            {
                // -----------------------------------------------------------------
                // 3. Render each chart shape to a high‑resolution PNG.
                // -----------------------------------------------------------------
                ShapeRenderer renderer = shape.GetShapeRenderer();

                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    // 300 DPI yields a high‑resolution image.
                    Resolution = 300
                };

                string imagePath = Path.Combine(artifactsDir, $"ChartImage.{chartCount}.png");
                renderer.Save(imagePath, saveOptions);
                chartCount++;
            }
        }

        // Validate that at least one chart image was extracted.
        if (chartCount == 0)
            throw new InvalidOperationException("No chart objects were found in the document.");
    }
}
