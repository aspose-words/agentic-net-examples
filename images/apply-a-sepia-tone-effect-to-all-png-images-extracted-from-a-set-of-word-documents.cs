using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class SepiaImageProcessor
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "SepiaImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample PNG image
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        CreateSamplePng(sampleImagePath);

        // Create sample Word documents containing the PNG image
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // Process each document: extract PNG images, apply sepia, save result
        int extractedCount = 0;
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int shapeIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    // Save the image to a memory stream
                    using (MemoryStream ms = new MemoryStream())
                    {
                        shape.ImageData.Save(ms);
                        ms.Position = 0; // Reset stream position

                        // Load bitmap from stream
                        using (Bitmap bitmap = new Bitmap(ms))
                        {
                            // Apply sepia tone
                            ApplySepia(bitmap);

                            // Save the processed image
                            string outFileName = $"{Path.GetFileNameWithoutExtension(docFile)}-shape{shapeIndex}.png";
                            string outPath = Path.Combine(outputFolder, outFileName);
                            bitmap.Save(outPath, ImageFormat.Png);
                            extractedCount++;
                        }
                    }
                    shapeIndex++;
                }
            }
        }

        // Validate that at least one image was processed
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No PNG images were extracted and processed.");
        }
    }

    private static void CreateSamplePng(string path)
    {
        int width = 100;
        int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle
                using (SolidBrush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 10, 10, 80, 80);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }

    private static void ApplySepia(Bitmap bitmap)
    {
        int width = bitmap.Width;
        int height = bitmap.Height;

        for (int y = 0; y < height; y++)
        {
            for (int x = 0; x < width; x++)
            {
                Color original = bitmap.GetPixel(x, y);
                double r = original.R;
                double g = original.G;
                double b = original.B;

                int tr = (int)(0.393 * r + 0.769 * g + 0.189 * b);
                int tg = (int)(0.349 * r + 0.686 * g + 0.168 * b);
                int tb = (int)(0.272 * r + 0.534 * g + 0.131 * b);

                tr = Math.Min(255, tr);
                tg = Math.Min(255, tg);
                tb = Math.Min(255, tb);

                Color sepia = Color.FromArgb(tr, tg, tb);
                bitmap.SetPixel(x, y, sepia);
            }
        }
    }
}
