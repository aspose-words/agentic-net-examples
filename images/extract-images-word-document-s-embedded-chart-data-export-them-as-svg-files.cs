using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;

class ExtractChartImagesToSvg
{
    static void Main()
    {
        // Path to the source Word document.
        string sourceDocPath = @"C:\Docs\InputDocument.docx";

        // Folder where the extracted SVG files will be saved.
        string svgOutputFolder = @"C:\Docs\ChartSvgs";

        // Ensure the output folder exists.
        Directory.CreateDirectory(svgOutputFolder);

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Get all shape nodes in the document (including charts).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int chartIndex = 0;

        // Iterate through each shape and process only those that contain a chart.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasChart)
            {
                // Build the SVG file name.
                string svgFilePath = Path.Combine(svgOutputFolder, $"Chart_{chartIndex}.svg");

                // Configure SVG save options.
                SvgSaveOptions svgOptions = new SvgSaveOptions
                {
                    // Render text as curves to avoid font issues.
                    TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
                    // Do not embed raster images inside the SVG (charts are vector based).
                    ExportEmbeddedImages = false,
                    // Remove the page border for a cleaner output.
                    ShowPageBorder = false,
                    // Fit the SVG to the viewport (optional).
                    FitToViewPort = true
                };

                // Render the chart shape to an SVG file.
                shape.GetShapeRenderer().Save(svgFilePath, svgOptions);

                chartIndex++;
            }
        }

        Console.WriteLine($"Extracted {chartIndex} chart(s) to SVG files in \"{svgOutputFolder}\".");
    }
}
