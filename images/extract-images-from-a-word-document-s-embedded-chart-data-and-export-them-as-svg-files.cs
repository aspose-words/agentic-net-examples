using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class ExtractChartSvg
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and insert a sample column chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart. The returned Shape contains the chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Populate the chart with sample data.
        string[] categories = { "Category A", "Category B", "Category C" };
        double[] values = { 10, 20, 30 };
        chart.Series.Add("Series 1", categories, values);

        // Find all shapes that contain charts.
        var chartShapes = doc.GetChildNodes(NodeType.Shape, true)
                             .OfType<Shape>()
                             .Where(s => s.HasChart)
                             .ToList();

        if (!chartShapes.Any())
            throw new InvalidOperationException("No chart found in the document.");

        int chartIndex = 0;
        foreach (Shape shape in chartShapes)
        {
            // Render the chart shape to an SVG file.
            var svgOptions = new SvgSaveOptions
            {
                ExportEmbeddedImages = false,
                ResourcesFolder = outputDir
            };

            string svgPath = Path.Combine(outputDir, $"Chart_{chartIndex}.svg");
            shape.GetShapeRenderer().Save(svgPath, svgOptions);

            // Validate that the SVG file was created.
            if (!File.Exists(svgPath))
                throw new FileNotFoundException($"Failed to create SVG file: {svgPath}");

            chartIndex++;
        }

        // Save the original document for reference.
        string docPath = Path.Combine(outputDir, "DocumentWithChart.docx");
        doc.Save(docPath);
    }
}
