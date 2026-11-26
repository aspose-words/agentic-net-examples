using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define input and output folders
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample BMP images if none exist
        CreateSampleBmpImages(inputFolder, 3);

        // Get JPEG encoder
        ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
            .FirstOrDefault(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
        if (jpegCodec == null)
            throw new InvalidOperationException("JPEG codec not found.");

        // Prepare encoder parameters for 80% quality
        EncoderParameters encoderParams = new EncoderParameters(1);
        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 80L);

        // Process each BMP file
        string[] bmpFiles = Directory.GetFiles(inputFolder, "*.bmp");
        if (bmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP files found to convert.");

        int successCount = 0;
        foreach (string bmpPath in bmpFiles)
        {
            try
            {
                using (Bitmap bitmap = new Bitmap(bmpPath))
                {
                    string outputFileName = Path.GetFileNameWithoutExtension(bmpPath) + ".jpg";
                    string outputPath = Path.Combine(outputFolder, outputFileName);
                    bitmap.Save(outputPath, jpegCodec, encoderParams);
                    Console.WriteLine($"Converted '{Path.GetFileName(bmpPath)}' to '{outputFileName}' successfully.");
                    successCount++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to convert '{Path.GetFileName(bmpPath)}': {ex.Message}");
            }
        }

        // Validate that at least one JPEG was produced
        if (successCount == 0)
            throw new InvalidOperationException("No BMP images were successfully converted to JPEG.");

        Console.WriteLine($"Conversion completed. {successCount} file(s) converted.");
    }

    // Helper method to create deterministic sample BMP images
    private static void CreateSampleBmpImages(string folderPath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            string filePath = Path.Combine(folderPath, $"sample{i}.bmp");
            if (File.Exists(filePath))
                continue; // Skip if already exists

            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.White);
                    // Draw a simple rectangle with a distinct color
                    using (Pen pen = new Pen(Color.FromArgb(255, 0, 0, 255), 5))
                    {
                        g.DrawRectangle(pen, 20, 20, 160, 160);
                    }
                }
                bitmap.Save(filePath, ImageFormat.Bmp);
            }
        }
    }
}
