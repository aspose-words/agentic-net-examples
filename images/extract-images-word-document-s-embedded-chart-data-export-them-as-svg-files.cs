using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

class ExtractChartImagesToSvg
{
    static void Main()
    {
        // Output folder for the extracted SVG files (relative to the executable directory).
        string svgOutputFolder = Path.Combine(Environment.CurrentDirectory, "ChartSvgs");
        Directory.CreateDirectory(svgOutputFolder);

        // Create a sample Word document with a chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertChart(ChartType.Column, 400, 300);

        // Get all shape nodes in the document (including charts).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int chartIndex = 0;

        // Iterate through each shape and process only those that contain a chart.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasChart)
            {
                string svgFilePath = Path.Combine(svgOutputFolder, $"Chart_{chartIndex}.svg");

                var svgOptions = new SvgSaveOptions
                {
                    TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
                    ExportEmbeddedImages = false,
                    ShowPageBorder = false,
                    FitToViewPort = true
                };

                shape.GetShapeRenderer().Save(svgFilePath, svgOptions);
                chartIndex++;
            }
        }

        Console.WriteLine($"Extracted {chartIndex} chart(s) to SVG files in \"{svgOutputFolder}\".");
    }
}
