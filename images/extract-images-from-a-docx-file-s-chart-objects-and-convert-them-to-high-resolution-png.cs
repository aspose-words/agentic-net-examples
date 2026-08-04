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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX containing a chart.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "ChartDocument.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Populate the chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Series 1", new[] { "A", "B", "C" }, new[] { 10.0, 20.0, 30.0 });
        chart.Series.Add("Series 2", new[] { "A", "B", "C" }, new[] { 15.0, 25.0, 35.0 });

        // Save the document.
        doc.Save(docPath);
        // -----------------------------------------------------------------

        // -----------------------------------------------------------------
        // 2. Load the document and extract chart shapes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int chartIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Chart shapes expose a non‑null Chart property.
            if (shape.Chart != null)
            {
                // Render the chart shape to a high‑resolution PNG.
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    // Set a high DPI for better quality.
                    Resolution = 300f
                };

                string imagePath = Path.Combine(artifactsDir, $"ChartImage_{chartIndex}.png");
                shape.GetShapeRenderer().Save(imagePath, options);
                chartIndex++;
            }
        }
        // -----------------------------------------------------------------

        // Validate that at least one image was extracted.
        if (chartIndex == 0)
            throw new InvalidOperationException("No chart images were extracted from the document.");

        Console.WriteLine($"Extracted {chartIndex} chart image(s) to folder: {artifactsDir}");
    }
}
