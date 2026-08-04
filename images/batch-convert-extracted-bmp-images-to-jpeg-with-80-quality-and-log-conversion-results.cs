using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchBmpToJpegConverter
{
    public static void Main()
    {
        // Define folders for input BMPs and output JPEGs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");

        // Ensure folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample BMP images.
        CreateSampleBmp(Path.Combine(inputFolder, "sample1.bmp"), 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleBmp(Path.Combine(inputFolder, "sample2.bmp"), 300, 200, Aspose.Drawing.Color.LightGreen);
        CreateSampleBmp(Path.Combine(inputFolder, "sample3.bmp"), 250, 250, Aspose.Drawing.Color.LightCoral);

        // Prepare conversion options: JPEG with 80% quality.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 80
        };

        int convertedCount = 0;

        // Process each BMP file in the input folder.
        foreach (string bmpPath in Directory.GetFiles(inputFolder, "*.bmp"))
        {
            // Load the BMP into a new document and insert it as an image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(bmpPath);

            // Determine output JPEG path.
            string outputFileName = Path.GetFileNameWithoutExtension(bmpPath) + ".jpg";
            string jpegPath = Path.Combine(outputFolder, outputFileName);

            // Save the document page (containing the image) as a JPEG.
            doc.Save(jpegPath, jpegOptions);

            // Verify that the JPEG file was created.
            if (!File.Exists(jpegPath))
                throw new InvalidOperationException($"Failed to create JPEG file: {jpegPath}");

            // Log the conversion result.
            Console.WriteLine($"Converted '{Path.GetFileName(bmpPath)}' to '{outputFileName}' with JPEG quality {jpegOptions.JpegQuality}.");
            convertedCount++;
        }

        // Final summary.
        Console.WriteLine($"Batch conversion completed. Total files converted: {convertedCount}.");
    }

    // Helper method to create a deterministic BMP image.
    private static void CreateSampleBmp(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        // Create a bitmap with the specified dimensions.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain a graphics object to draw on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with a solid background color.
                graphics.Clear(backgroundColor);
            }

            // Save the bitmap as a BMP file.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Bmp);
        }

        // Validate that the BMP file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create BMP file: {filePath}");
    }
}
