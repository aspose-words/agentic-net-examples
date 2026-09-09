using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchTiffToJpegConverter
{
    public static void Main()
    {
        // Define folders for input TIFF images and output JPEG images.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");

        // Ensure clean directories.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample TIFF images.
        CreateSampleTiffImages(inputFolder, 3);

        // Batch convert each TIFF to JPEG with 90% quality.
        var tiffFiles = Directory.GetFiles(inputFolder, "*.*")
                                 .Where(f => f.EndsWith(".tif", StringComparison.OrdinalIgnoreCase) ||
                                             f.EndsWith(".tiff", StringComparison.OrdinalIgnoreCase))
                                 .ToArray();

        foreach (string tiffPath in tiffFiles)
        {
            // Load the TIFF into a new document by inserting it as an image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(tiffPath);

            // Configure JPEG save options with 90% quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 90
            };

            // Determine output file name.
            string outputFileName = Path.GetFileNameWithoutExtension(tiffPath) + ".jpg";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the rendered page as JPEG.
            doc.Save(outputPath, jpegOptions);
        }

        // Validate that at least one JPEG was produced.
        int jpegCount = Directory.GetFiles(outputFolder, "*.jpg").Length;
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG files were created during conversion.");

        // Example completed.
    }

    private static void CreateSampleTiffImages(string folderPath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            // Create a 200x200 bitmap.
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                // Fill with white background.
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);
                    // Draw a simple rectangle with a distinct color.
                    graphics.FillRectangle(new SolidBrush(Color.FromArgb(255, 100 + i * 30, 150, 200)), 20, 20, 160, 160);
                }

                // Save as TIFF.
                string filePath = Path.Combine(folderPath, $"Sample{i}.tiff");
                bitmap.Save(filePath, ImageFormat.Tiff);
            }
        }
    }
}
