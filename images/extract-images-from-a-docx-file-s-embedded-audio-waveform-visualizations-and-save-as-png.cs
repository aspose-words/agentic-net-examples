using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample waveform image.
        string waveformPath = Path.Combine(artifactsDir, "waveform.png");
        CreateWaveformImage(waveformPath, 300, 100);

        // Create a DOCX and insert the waveform image.
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, waveformPath);

        // Load the document and extract all images, saving them as PNG files.
        ExtractImagesAsPng(docPath, artifactsDir);
    }

    private static void CreateWaveformImage(string filePath, int width, int height)
    {
        // Create a bitmap and draw a simple sine‑wave like pattern.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 2))
            {
                for (int x = 0; x < width; x++)
                {
                    double radians = (double)x / width * 4 * Math.PI;
                    int y = (int)(height / 2 + Math.Sin(radians) * height / 3);
                    if (x > 0)
                    {
                        double prevRadians = (double)(x - 1) / width * 4 * Math.PI;
                        int prevY = (int)(height / 2 + Math.Sin(prevRadians) * height / 3);
                        g.DrawLine(pen, x - 1, prevY, x, y);
                    }
                }
            }
            bitmap.Save(filePath);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image into the document.
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ExtractImagesAsPng(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Force PNG extension regardless of original type.
                string outFile = Path.Combine(outputDir, $"extracted_{imageIndex}.png");
                shape.ImageData.Save(outFile);
                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
