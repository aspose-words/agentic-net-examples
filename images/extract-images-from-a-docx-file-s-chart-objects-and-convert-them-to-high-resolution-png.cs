using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ExtractChartImages
{
    public static void Main()
    {
        // Define deterministic folders and file names.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "SampleWithChart.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that contains a chart.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Populate the chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Series 1",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 10.0, 20.0, 30.0 });
        chart.Series.Add("Series 2",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 15.0, 25.0, 35.0 });

        // Save the document so that it can be re‑loaded later.
        doc.Save(docPath);
        // -----------------------------------------------------------------

        // -----------------------------------------------------------------
        // 2. Load the document and extract each chart as a high‑resolution PNG.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int chartIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // A shape that contains a chart has a non‑null Chart property.
            if (shape.Chart != null)
            {
                // Render the chart shape to a PNG image with high resolution (300 dpi).
                ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    Resolution = 300 // high‑resolution output
                };

                string pngPath = Path.Combine(
                    artifactsDir,
                    $"ChartImage.{chartIndex}{FileFormatUtil.ImageTypeToExtension(ImageType.Png)}");

                shape.GetShapeRenderer().Save(pngPath, pngOptions);
                chartIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 3. Validate that at least one chart image was extracted.
        // -----------------------------------------------------------------
        if (chartIndex == 0)
            throw new InvalidOperationException("No chart images were extracted from the document.");
    }
}
