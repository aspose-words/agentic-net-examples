using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class ExtractChartImagesToSvg
{
    public static void Main()
    {
        // Define folders for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "SampleChart.docx");

        // -------------------------------------------------
        // 1. Create a Word document that contains a chart.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;
        chart.Title.Text = "Sample Data";

        // Add a series with categories and corresponding values.
        chart.Series.Add(
            "Series 1",
            new[] { "A", "B", "C" },
            new double[] { 10, 20, 30 });

        // Save the document (demonstrates the create‑save lifecycle).
        doc.Save(docPath);

        // -------------------------------------------------
        // 2. Load the document and extract each chart as SVG.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int chartIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Identify chart shapes using the HasChart property.
            if (shape.HasChart)
            {
                string svgFileName = Path.Combine(artifactsDir, $"Chart_{chartIndex}.svg");

                // Render the chart shape to SVG.
                SvgSaveOptions svgOptions = new SvgSaveOptions
                {
                    ExportEmbeddedImages = false, // keep images separate if any.
                    ShowPageBorder = false
                };
                shape.GetShapeRenderer().Save(svgFileName, svgOptions);
                chartIndex++;
            }
        }

        // -------------------------------------------------
        // 3. Validate that at least one SVG file was created.
        // -------------------------------------------------
        if (chartIndex == 0)
            throw new InvalidOperationException("No chart shapes were found in the document.");

        // List the generated SVG files.
        string[] generatedSvgs = Directory.GetFiles(artifactsDir, "Chart_*.svg");
        foreach (string file in generatedSvgs)
            Console.WriteLine($"Generated SVG: {file}");
    }
}
