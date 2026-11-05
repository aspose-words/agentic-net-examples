using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files and output folder
        string inputImagePath = "waveform.png";
        string docPath = "sample.docx";
        string outputDir = "ExtractedImages";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample waveform image
        CreateSampleWaveformImage(inputImagePath);

        // Create a new document and insert the sample image (simulating an audio waveform visualization)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Load the document (simulating extraction from an existing file)
        Document loadedDoc = new Document(docPath);

        // Extract all images from shape nodes and save them as PNG files
        int imageCount = 0;
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outPath = Path.Combine(outputDir, $"extracted-{imageCount}.png");
                shape.ImageData.Save(outPath);
                imageCount++;
            }
        }

        // Validate that at least one image was extracted
        if (imageCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    private static void CreateSampleWaveformImage(string filePath)
    {
        int width = 200;
        int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);

                // Draw a simple waveform using vertical lines
                using (Pen pen = new Pen(Color.Black, 2))
                {
                    for (int x = 0; x < width; x += 10)
                    {
                        double radians = (double)x / width * 2 * Math.PI;
                        int y = (int)(height / 2 + (height / 4) * Math.Sin(radians));
                        graphics.DrawLine(pen, x, height / 2, x, y);
                    }
                }
            }

            bitmap.Save(filePath);
        }
    }
}
