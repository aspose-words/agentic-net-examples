using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class ExtractChartImagesAsSvg
{
    public static void Main()
    {
        // Define a deterministic folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "ChartDocument.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains a chart.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart. InsertChart returns a Shape that holds the chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart; // Access the Chart object from the shape.

        chart.Title.Text = "Sample Chart";

        // Populate the chart with two series.
        chart.Series.Clear();
        chart.Series.Add("Series 1",
            new string[] { "Category A", "Category B", "Category C" },
            new double[] { 10, 20, 30 });
        chart.Series.Add("Series 2",
            new string[] { "Category A", "Category B", "Category C" },
            new double[] { 15, 25, 35 });

        // Save the document containing the chart.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract chart shapes as SVG files.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Find all shapes that contain a chart.
        var chartShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .Where(s => s.HasChart)
                                   .ToList();

        if (!chartShapes.Any())
            throw new InvalidOperationException("No chart shapes were found in the document.");

        int chartIndex = 0;
        foreach (var shape in chartShapes)
        {
            // Configure SVG save options – do not embed raster images.
            SvgSaveOptions svgOptions = new SvgSaveOptions
            {
                ExportEmbeddedImages = false,
                ResourcesFolder = artifactsDir,
                ResourcesFolderAlias = artifactsDir,
                ShowPageBorder = false
            };

            // Build a deterministic SVG file name.
            string svgFileName = Path.Combine(artifactsDir,
                $"Chart_{chartIndex}.svg");

            // Render the chart shape to an SVG file.
            shape.GetShapeRenderer().Save(svgFileName, svgOptions);
            chartIndex++;
        }

        // Validate that at least one SVG file was created.
        var svgFiles = Directory.GetFiles(artifactsDir, "*.svg");
        if (!svgFiles.Any())
            throw new InvalidOperationException("No SVG files were generated from chart shapes.");

        Console.WriteLine($"Extracted {svgFiles.Length} SVG file(s) to \"{artifactsDir}\".");
    }
}
