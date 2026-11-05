using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "InputDocs";
        string outputFolder = "OutputImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample PNG image
        string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath);

        // Create sample Word documents containing the PNG image
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), sampleImagePath);

        // Process each document: extract PNG images, apply sepia, save result
        int totalSepiaImages = 0;
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (!shape.HasImage)
                    continue;

                if (shape.ImageData.ImageType != ImageType.Png)
                    continue;

                // Save the original image to a memory stream
                using (MemoryStream imageStream = new MemoryStream())
                {
                    shape.ImageData.Save(imageStream);
                    imageStream.Position = 0; // Reset before reading

                    // Load bitmap from stream
                    using (Bitmap bitmap = new Bitmap(imageStream))
                    {
                        ApplySepiaTone(bitmap);

                        // Save sepia image to output folder
                        string outputPath = Path.Combine(
                            outputFolder,
                            $"sepia-{Path.GetFileNameWithoutExtension(docPath)}-{imageIndex}.png");

                        bitmap.Save(outputPath, ImageFormat.Png);
                        totalSepiaImages++;
                    }
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one sepia image was created
        if (totalSepiaImages == 0)
            throw new InvalidOperationException("No PNG images were processed and saved.");

        // Cleanup sample files (optional)
        // File.Delete(sampleImagePath);
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

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private static void ApplySepiaTone(Bitmap bitmap)
    {
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
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
