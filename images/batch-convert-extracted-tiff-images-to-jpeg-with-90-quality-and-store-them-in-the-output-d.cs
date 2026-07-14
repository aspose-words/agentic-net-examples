using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchTiffToJpegConverter
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample TIFF images.
        CreateSampleTiff(Path.Combine(inputDir, "sample1.tif"), 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleTiff(Path.Combine(inputDir, "sample2.tif"), 300, 200, Aspose.Drawing.Color.LightCoral);

        // Build a document and insert the TIFF images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string tiffPath in Directory.GetFiles(inputDir, "*.tif"))
        {
            builder.InsertParagraph();
            builder.InsertImage(tiffPath);
        }

        // Extract each image from the document and save it as JPEG with 90% quality.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int jpegIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string jpegPath = Path.Combine(outputDir, $"converted_{jpegIndex}.jpg");
                ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
                {
                    JpegQuality = 90
                };
                // Render the shape to JPEG.
                shape.GetShapeRenderer().Save(jpegPath, jpegOptions);
                jpegIndex++;
            }
        }

        // Validate that at least one JPEG file was produced.
        if (jpegIndex == 0)
            throw new InvalidOperationException("No images were found for conversion.");
    }

    // Generates a deterministic TIFF image using Aspose.Drawing.
    private static void CreateSampleTiff(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
            }
            bitmap.Save(filePath, ImageFormat.Tiff);
        }
    }
}
