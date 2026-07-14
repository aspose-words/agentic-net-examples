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
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with a chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Populate chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Series 1",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 10.0, 20.0, 30.0 });
        chart.Series.Add("Series 2",
            new[] { "Category A", "Category B", "Category C" },
            new[] { 15.0, 25.0, 35.0 });

        // Save the document (creation step).
        string docPath = Path.Combine(outputDir, "SampleWithChart.docx");
        doc.Save(docPath);

        // Load the document (loading step).
        Document loadedDoc = new Document(docPath);

        // Find all chart shapes.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int svgIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Chart shapes are OLE objects; they expose HasChart.
            if (shape.HasChart)
            {
                // Render the chart shape to an SVG file.
                string svgPath = Path.Combine(outputDir, $"ChartImage_{svgIndex}.svg");
                SvgSaveOptions svgOptions = new SvgSaveOptions
                {
                    ExportEmbeddedImages = false,
                    ResourcesFolder = outputDir,
                    ShowPageBorder = false
                };
                shape.GetShapeRenderer().Save(svgPath, svgOptions);
                svgIndex++;
            }
        }

        // Validate that at least one SVG file was created.
        if (svgIndex == 0)
            throw new InvalidOperationException("No chart shapes were found to export as SVG.");

        // List generated SVG files.
        foreach (string file in Directory.GetFiles(outputDir, "*.svg"))
            Console.WriteLine($"Generated SVG: {file}");
    }
}
