using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing provides Bitmap, Graphics, Color, Pen

public class BatchBmpToWebp
{
    public static void Main()
    {
        // Prepare deterministic folders for input and output images.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample BMP images that will be used as the batch source.
        CreateSampleBmp(Path.Combine(inputDir, "sample1.bmp"), 200, 150, Color.LightBlue);
        CreateSampleBmp(Path.Combine(inputDir, "sample2.bmp"), 300, 200, Color.LightGreen);
        CreateSampleBmp(Path.Combine(inputDir, "sample3.bmp"), 250, 250, Color.LightCoral);

        // Process each BMP file in the input folder.
        foreach (string bmpPath in Directory.GetFiles(inputDir, "*.bmp"))
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(bmpPath);
            string webpPath = Path.Combine(outputDir, fileNameWithoutExt + ".webp");

            // Load the BMP into a temporary document and insert the image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(bmpPath);

            // Save the single‑page document as a WebP image (lossless is default for WebP).
            ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
            doc.Save(webpPath, webpOptions);

            // Log conversion details.
            FileInfo originalInfo = new FileInfo(bmpPath);
            FileInfo convertedInfo = new FileInfo(webpPath);
            Console.WriteLine($"Converted '{originalInfo.Name}' ({originalInfo.Length} bytes) to '{convertedInfo.Name}' ({convertedInfo.Length} bytes).");
        }

        // Verify that at least one WebP file was created.
        if (Directory.GetFiles(outputDir, "*.webp").Length == 0)
            throw new InvalidOperationException("No WebP files were created during conversion.");
    }

    // Creates a deterministic BMP file using Aspose.Drawing.
    private static void CreateSampleBmp(string filePath, int width, int height, Color backgroundColor)
    {
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(backgroundColor);

        // Draw a simple rectangle for visual distinction.
        using (Pen pen = new Pen(Color.Black, 3))
        {
            graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
        }

        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }
}
