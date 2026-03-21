using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class ChartImageExtractor
{
    /// <summary>
    /// Extracts all chart objects from a DOCX file and saves them as high‑resolution PNG images.
    /// </summary>
    /// <param name="docPath">Full path to the source DOCX document.</param>
    /// <param name="outputDir">Folder where the PNG files will be written.</param>
    public static void ExtractChartImages(string docPath, string outputDir)
    {
        // Ensure the output directory exists.
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Load the Word document.
        Document doc = new Document(docPath);

        // Collect all Shape nodes that contain a chart.
        var chartShapes = doc.GetChildNodes(NodeType.Shape, true)
                             .OfType<Shape>()
                             .Where(s => s.HasChart);

        int chartIndex = 0;
        foreach (Shape chartShape in chartShapes)
        {
            // Prepare high‑resolution PNG save options.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Set a high DPI (e.g., 300) for better quality.
                Resolution = 300,
                // Optional: increase the scale factor if you need even larger images.
                Scale = 1.0f
            };

            // Build the output file name.
            string pngPath = Path.Combine(outputDir, $"Chart_{chartIndex}.png");

            // Render the chart shape to PNG.
            chartShape.GetShapeRenderer().Save(pngPath, pngOptions);

            chartIndex++;
        }

        Console.WriteLine($"{chartIndex} chart image(s) extracted to \"{outputDir}\".");
    }

    // Example usage.
    public static void Main()
    {
        string sourceDoc = @"C:\Docs\SampleWithCharts.docx";
        string imagesFolder = @"C:\Docs\ExtractedCharts";

        if (!File.Exists(sourceDoc))
        {
            Console.WriteLine($"Source file not found: {sourceDoc}");
            Console.WriteLine("Please place a DOCX file with charts at the specified location or update the path.");
            return;
        }

        ExtractChartImages(sourceDoc, imagesFolder);
        Console.WriteLine("Chart images extraction completed.");
    }
}
