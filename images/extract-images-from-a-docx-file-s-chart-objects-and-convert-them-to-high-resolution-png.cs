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
        // Folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that contains a chart.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart and obtain the underlying Chart object.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Populate the chart with some data using the overload that accepts
        // categories and values directly.
        chart.Series.Clear();
        chart.Series.Add("Series 1", new[] { "A", "B", "C" }, new[] { 10.0, 20.0, 30.0 });

        // Save the sample document (optional, just for reference).
        string docPath = Path.Combine(artifactsDir, "SampleWithChart.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Extract each chart as a high‑resolution PNG image.
        // -----------------------------------------------------------------
        int chartIndex = 0;
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Identify shapes that contain a chart.
            if (shape.Chart != null)
            {
                // Render the chart to a PNG with high DPI (e.g., 300).
                var renderer = shape.GetShapeRenderer();
                var saveOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    Resolution = 300 // high‑resolution DPI
                };

                string imageFile = Path.Combine(
                    artifactsDir,
                    $"ChartImage_{chartIndex}{FileFormatUtil.ImageTypeToExtension(ImageType.Png)}");

                renderer.Save(imageFile, saveOptions);
                chartIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 3. Validation.
        // -----------------------------------------------------------------
        if (chartIndex == 0)
            throw new InvalidOperationException("No chart images were extracted.");

        Console.WriteLine($"{chartIndex} chart image(s) extracted to \"{artifactsDir}\".");
    }
}
