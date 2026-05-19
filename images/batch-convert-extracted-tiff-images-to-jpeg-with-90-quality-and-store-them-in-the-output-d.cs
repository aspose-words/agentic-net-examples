using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchTiffToJpegConverter
{
    public static void Main()
    {
        // Define folders relative to the executable location.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");

        // Ensure a clean state for the demo.
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create deterministic sample TIFF images.
        CreateSampleTiff(Path.Combine(inputDir, "sample1.tiff"));
        CreateSampleTiff(Path.Combine(inputDir, "sample2.tiff"));

        // Prepare JPEG save options with 90% quality.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 90
        };

        // Process each TIFF file in the input folder.
        string[] tiffFiles = Directory.GetFiles(inputDir, "*.tiff");
        int processedCount = 0;

        foreach (string tiffPath in tiffFiles)
        {
            // Load the TIFF into a new document and insert the image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(tiffPath);

            // Determine the output JPEG file name.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(tiffPath);
            string jpegPath = Path.Combine(outputDir, fileNameWithoutExt + ".jpg");

            // Save the single-page document as a JPEG image.
            doc.Save(jpegPath, jpegOptions);
            processedCount++;
        }

        // Validate that at least one JPEG was produced.
        if (processedCount == 0 || Directory.GetFiles(outputDir, "*.jpg").Length == 0)
            throw new InvalidOperationException("No JPEG files were produced.");

        // Example completed without interactive prompts.
    }

    // Creates a simple deterministic TIFF image using Aspose.Drawing.
    private static void CreateSampleTiff(string filePath)
    {
        int width = 200;
        int height = 200;

        // Create a bitmap and draw a light‑blue rectangle on a white background.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                using (SolidBrush brush = new SolidBrush(Color.LightBlue))
                {
                    g.FillRectangle(brush, 20, 20, width - 40, height - 40);
                }
            }

            // Save the bitmap as a TIFF file.
            bitmap.Save(filePath, ImageFormat.Tiff);
        }
    }
}
