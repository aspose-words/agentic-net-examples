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
        // Prepare folders
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample BMP images
        for (int i = 1; i <= 3; i++)
        {
            string bmpPath = Path.Combine(inputFolder, $"image{i}.bmp");
            using (Bitmap bmp = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    // Fill background
                    g.Clear(Aspose.Drawing.Color.White);
                    // Draw a simple rectangle with a distinct color
                    using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.FromArgb(50 * i, 100, 150)))
                    {
                        g.FillRectangle(brush, 20, 20, 160, 160);
                    }
                }
                bmp.Save(bmpPath, ImageFormat.Bmp);
            }
        }

        // Get JPEG encoder
        ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
            .FirstOrDefault(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
        if (jpegCodec == null)
            throw new InvalidOperationException("JPEG encoder not found.");

        // Prepare encoder parameters for 80% quality
        EncoderParameters encoderParams = new EncoderParameters(1);
        EncoderParameter qualityParam = new EncoderParameter(Encoder.Quality, 80L);
        encoderParams.Param[0] = qualityParam;

        // Process each BMP file
        string[] bmpFiles = Directory.GetFiles(inputFolder, "*.bmp");
        if (bmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP files found to convert.");

        foreach (string bmpFile in bmpFiles)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(bmpFile);
            string jpegPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.jpg");

            using (Bitmap bmp = new Bitmap(bmpFile))
            {
                // Save as JPEG with quality setting
                bmp.Save(jpegPath, jpegCodec, encoderParams);
            }

            Console.WriteLine($"Converted '{Path.GetFileName(bmpFile)}' to '{Path.GetFileName(jpegPath)}' with 80% quality.");
        }

        // Validate output
        int jpegCount = Directory.GetFiles(outputFolder, "*.jpg").Length;
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG files were created.");

        Console.WriteLine($"Batch conversion completed. {jpegCount} JPEG file(s) created in '{outputFolder}'.");
    }
}
