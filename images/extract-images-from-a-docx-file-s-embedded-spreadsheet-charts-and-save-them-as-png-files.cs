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
        // Create a sample DOCX document with an embedded chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple column chart (width = 432 points, height = 288 points).
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 288);
        Chart chart = chartShape.Chart;

        // Populate the chart with sample data.
        chart.Series.Clear();
        chart.Series.Add("Series 1",
            new string[] { "Category A", "Category B", "Category C" },
            new double[] { 10, 20, 30 });

        // Save the document to a local file.
        const string docPath = "SampleWithChart.docx";
        doc.Save(docPath);

        // Reload the document (demonstrates load/save lifecycle).
        Document loadedDoc = new Document(docPath);

        // Prepare output folder.
        const string outputFolder = "ExtractedImages";
        Directory.CreateDirectory(outputFolder);

        // Iterate through all Shape nodes and extract images or chart renderings.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // ----- Extract embedded raster images -----
            if (shape.HasImage)
            {
                // Determine proper file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imagePath = Path.Combine(outputFolder, $"ChartImage_{imageIndex}{extension}");

                // Save the image directly from ImageData.
                shape.ImageData.Save(imagePath);

                // Validate that the file was created.
                if (!File.Exists(imagePath))
                    throw new InvalidOperationException($"Failed to save image: {imagePath}");

                Console.WriteLine($"Extracted image saved to: {imagePath}");
                imageIndex++;
                continue;
            }

            // ----- Extract chart as PNG image -----
            if (shape.Chart != null)
            {
                // Render the chart shape to a PNG file using ShapeRenderer.
                string pngPath = Path.Combine(outputFolder, $"ChartImage_{imageIndex}.png");
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
                shape.GetShapeRenderer().Save(pngPath, saveOptions);

                // Validate that the file was created.
                if (!File.Exists(pngPath))
                    throw new InvalidOperationException($"Failed to save chart image: {pngPath}");

                Console.WriteLine($"Extracted chart saved to: {pngPath}");
                imageIndex++;
            }
        }

        // Ensure at least one image or chart was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images or charts were found in the document.");

        // Optional cleanup of the temporary DOCX file.
        // File.Delete(docPath);
    }
}
