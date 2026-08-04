using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchTiffToJpegConverter
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample TIFF images.
        CreateSampleTiffImages(inputDir, count: 3);

        // Convert each TIFF image to JPEG with 90% quality.
        int convertedCount = 0;
        foreach (string tiffPath in Directory.GetFiles(inputDir, "*.tiff"))
        {
            // Load a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the TIFF image into the document.
            builder.InsertImage(tiffPath);

            // Configure JPEG save options with 90% quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 90
            };

            // Determine output JPEG file path.
            string jpegFileName = Path.GetFileNameWithoutExtension(tiffPath) + ".jpg";
            string jpegPath = Path.Combine(outputDir, jpegFileName);

            // Save the document page (containing the image) as JPEG.
            doc.Save(jpegPath, jpegOptions);

            // Validate that the JPEG file was created.
            if (!File.Exists(jpegPath))
                throw new InvalidOperationException($"Failed to create JPEG file: {jpegPath}");

            convertedCount++;
        }

        // Ensure at least one image was converted.
        if (convertedCount == 0)
            throw new InvalidOperationException("No TIFF images were found for conversion.");
    }

    // Helper method to create deterministic sample TIFF images.
    private static void CreateSampleTiffImages(string folderPath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            string filePath = Path.Combine(folderPath, $"sample{i}.tiff");

            // Create a 200x200 bitmap.
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                // Fill the bitmap with a solid color.
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.FromArgb(255, 100 + i * 30, 150, 200));
                }

                // Save as TIFF.
                bitmap.Save(filePath);
            }

            // Validate that the TIFF file was created.
            if (!File.Exists(filePath))
                throw new InvalidOperationException($"Failed to create sample TIFF image: {filePath}");
        }
    }
}
