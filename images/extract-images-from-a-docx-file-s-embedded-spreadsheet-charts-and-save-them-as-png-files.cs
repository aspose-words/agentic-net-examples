using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;          // For ChartType enum
using Aspose.Words.Rendering;               // For ShapeRenderer
using Aspose.Words.Saving;                  // For ImageSaveOptions

public class ExtractChartImages
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare folders.
        // -----------------------------------------------------------------
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputDir = Path.Combine(artifactsDir, "ChartImages");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 2. Create a sample DOCX that contains an embedded spreadsheet chart.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart. The chart is stored as a shape.
        builder.InsertChart(ChartType.Column, 432, 288);
        string sampleDocPath = Path.Combine(artifactsDir, "SampleWithChart.docx");
        doc.Save(sampleDocPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images from embedded charts.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sampleDocPath);

        // Get all shape nodes (charts are stored as shapes).
        var chartShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>();

        int chartIndex = 0;
        foreach (Shape chartShape in chartShapes)
        {
            // Render the shape (chart) to a PNG image.
            ShapeRenderer renderer = chartShape.GetShapeRenderer();
            string imagePath = Path.Combine(outputDir, $"ChartImage_{chartIndex}.png");
            renderer.Save(imagePath, new ImageSaveOptions(SaveFormat.Png));
            chartIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // -----------------------------------------------------------------
        if (chartIndex == 0)
            throw new InvalidOperationException("No chart images were extracted from the document.");

        Console.WriteLine($"Extracted {chartIndex} chart image(s) to \"{outputDir}\".");
    }
}
