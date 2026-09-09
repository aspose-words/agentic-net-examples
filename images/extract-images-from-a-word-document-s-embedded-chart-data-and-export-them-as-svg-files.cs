using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with a chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart.
        Chart chart = builder.InsertChart(ChartType.Column, 432, 252).Chart;
        // Populate chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Series 1",
            new[] { "Category A", "Category B", "Category C" },
            new double[] { 10, 20, 30 });

        // Save the document (optional, just to have a file on disk).
        string docPath = Path.Combine(outputDir, "SampleChart.docx");
        doc.Save(docPath);

        // Reload the document to simulate a real extraction scenario.
        Document loadedDoc = new Document(docPath);

        // Find all chart shapes in the document.
        // Charts are stored as OLE objects; they expose the HasChart property.
        var chartShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .OfType<Shape>()
                                   .Where(s => s.HasChart)
                                   .ToList();

        if (!chartShapes.Any())
            throw new InvalidOperationException("No chart shapes were found in the document.");

        int chartIndex = 0;
        foreach (Shape chartShape in chartShapes)
        {
            // Render each chart shape to an SVG file.
            string svgFileName = Path.Combine(outputDir, $"Chart_{chartIndex}.svg");
            var svgOptions = new SvgSaveOptions
            {
                ExportEmbeddedImages = false,
                ShowPageBorder = false
            };
            chartShape.GetShapeRenderer().Save(svgFileName, svgOptions);
            chartIndex++;
        }

        // Validate that at least one SVG file was created.
        if (chartIndex == 0 || !Directory.EnumerateFiles(outputDir, "*.svg").Any())
            throw new InvalidOperationException("SVG extraction failed; no SVG files were created.");
    }
}
