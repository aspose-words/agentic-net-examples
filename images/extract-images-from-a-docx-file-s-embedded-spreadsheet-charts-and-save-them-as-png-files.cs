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
        // Prepare working folders.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string outputDir = Path.Combine(workDir, "ExtractedImages");
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that contains an embedded chart.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(workDir, "SampleWithChart.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart. The chart is stored as a Shape.
        builder.InsertChart(ChartType.Column, 400, 300);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract images from chart shapes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // If the shape already contains an image (e.g., a picture), save it directly.
            if (shape.HasImage)
            {
                string outFile = Path.Combine(outputDir,
                    $"ChartImage_{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
                shape.ImageData.Save(outFile);
                imageIndex++;
                continue;
            }

            // If the shape is a chart, render it to a PNG image.
            // In Aspose.Words the presence of a chart can be checked via HasChart.
            if (shape.HasChart)
            {
                string outFile = Path.Combine(outputDir, $"ChartImage_{imageIndex}.png");
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
                shape.GetShapeRenderer().Save(outFile, options);
                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 3. Validate that at least one image was extracted.
        // -----------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No chart images were extracted from the document.");
    }
}
