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
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample BMP images
        CreateSampleBmp(Path.Combine(inputFolder, "sample1.bmp"));
        CreateSampleBmp(Path.Combine(inputFolder, "sample2.bmp"));

        // Process each BMP file
        string[] bmpFiles = Directory.GetFiles(inputFolder, "*.bmp");
        int convertedCount = 0;

        foreach (string bmpPath in bmpFiles)
        {
            try
            {
                // Load image into a temporary document
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                Shape shape = builder.InsertImage(bmpPath);

                // Prepare output path with .webp extension
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(bmpPath);
                string outputPath = Path.Combine(outputFolder, fileNameWithoutExt + ".webp");

                // Save the image data as WebP (lossless if supported)
                shape.ImageData.Save(outputPath);

                // Log conversion details
                FileInfo inputInfo = new FileInfo(bmpPath);
                FileInfo outputInfo = new FileInfo(outputPath);
                Console.WriteLine($"Converted: '{inputInfo.Name}' ({inputInfo.Length} bytes) -> '{outputInfo.Name}' ({outputInfo.Length} bytes)");

                if (File.Exists(outputPath))
                    convertedCount++;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to convert '{Path.GetFileName(bmpPath)}': {ex.Message}");
            }
        }

        // Validation
        if (convertedCount == 0)
            throw new InvalidOperationException("No BMP images were converted to WebP.");

        Console.WriteLine($"Conversion completed. {convertedCount} file(s) created in '{outputFolder}'.");
    }

    private static void CreateSampleBmp(string filePath)
    {
        const int width = 200;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Bmp);
        }
    }
}
