using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class ExtractChartImages
{
    public static void Main()
    {
        // Define folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "SampleChart.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX containing a chart.
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
            new double[] { 10, 20, 30 });

        // Save the document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract each chart as a high‑resolution PNG.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int chartIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // A chart shape has a non‑null Chart property.
            if (shape.Chart != null)
            {
                // Render the chart shape to a PNG image with high resolution (300 DPI).
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    Resolution = 300,               // Sets both horizontal and vertical DPI.
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                string imagePath = Path.Combine(artifactsDir, $"ChartImage_{chartIndex}.png");
                shape.GetShapeRenderer().Save(imagePath, options);
                chartIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 3. Validate that at least one image was extracted.
        // -----------------------------------------------------------------
        if (chartIndex == 0)
            throw new InvalidOperationException("No chart images were extracted from the document.");
    }
}
