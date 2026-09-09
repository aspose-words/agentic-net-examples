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
        // Prepare output folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that contains an embedded Excel chart.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Set a title – data population is optional for extraction purposes.
        chart.Title.Text = "Sample Chart";

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "SampleWithChart.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract images from embedded charts.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Save any shape that already contains an image.
            if (shape.HasImage)
            {
                string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = Path.Combine(outputDir, $"Image_{imageIndex}{ext}");
                shape.ImageData.Save(fileName);
                imageIndex++;
                continue;
            }

            // If the shape is a chart, render it to PNG.
            if (shape.Chart != null)
            {
                string fileName = Path.Combine(outputDir, $"Chart_{imageIndex}.png");
                // Use ImageSaveOptions to specify PNG format.
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
                shape.GetShapeRenderer().Save(fileName, options);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (Directory.GetFiles(outputDir).Length == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
