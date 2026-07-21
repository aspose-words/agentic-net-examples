using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class SepiaImageProcessor
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(inputDocsDir);

        // Create a sample PNG image.
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSamplePng(sampleImagePath);

        // Create a few Word documents that contain the PNG image.
        const int documentCount = 2;
        for (int i = 1; i <= documentCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputDocsDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // Extract PNG images from each document and apply a sepia tone effect.
        int processedImageCount = 0;
        string[] docFiles = Directory.GetFiles(inputDocsDir, "*.docx");
        for (int docIndex = 0; docIndex < docFiles.Length; docIndex++)
        {
            Document doc = new Document(docFiles[docIndex]);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
                {
                    // Save the shape image to a memory stream.
                    using (MemoryStream imgStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imgStream);
                        imgStream.Position = 0; // Reset before reading.

                        // Load the image with Aspose.Drawing.
                        using (Bitmap bitmap = new Bitmap(imgStream))
                        {
                            ApplySepiaEffect(bitmap);
                            string outFile = Path.Combine(
                                artifactsDir,
                                $"sepia_doc{docIndex + 1}_img{imageIndex + 1}.png");
                            bitmap.Save(outFile);
                            processedImageCount++;
                        }
                    }

                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was processed.
        if (processedImageCount == 0)
            throw new InvalidOperationException("No PNG images were found and processed.");

        Console.WriteLine($"Processed {processedImageCount} PNG image(s).");
    }

    // Creates a deterministic sample PNG image.
    private static void CreateSamplePng(string filePath)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple blue rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(filePath);
        }
    }

    // Applies a sepia tone effect to the provided bitmap.
    private static void ApplySepiaEffect(Bitmap bitmap)
    {
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);
                double r = original.R;
                double g = original.G;
                double b = original.B;

                // Sepia conversion formula.
                int tr = (int)(0.393 * r + 0.769 * g + 0.189 * b);
                int tg = (int)(0.349 * r + 0.686 * g + 0.168 * b);
                int tb = (int)(0.272 * r + 0.534 * g + 0.131 * b);

                // Clamp values to [0,255].
                tr = Math.Min(255, tr);
                tg = Math.Min(255, tg);
                tb = Math.Min(255, tb);

                Aspose.Drawing.Color sepia = Aspose.Drawing.Color.FromArgb(original.A, tr, tg, tb);
                bitmap.SetPixel(x, y, sepia);
            }
        }
    }
}
