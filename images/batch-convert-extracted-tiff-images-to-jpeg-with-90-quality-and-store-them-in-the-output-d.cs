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
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "input");
        string outputDir = Path.Combine(baseDir, "output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample TIFF images
        CreateSampleTiff(Path.Combine(inputDir, "sample1.tif"));
        CreateSampleTiff(Path.Combine(inputDir, "sample2.tif"));

        // Get all TIFF files in the input folder
        string[] tiffFiles = Directory.GetFiles(inputDir, "*.*", SearchOption.TopDirectoryOnly)
                                      .Where(f => f.EndsWith(".tif", StringComparison.OrdinalIgnoreCase) ||
                                                  f.EndsWith(".tiff", StringComparison.OrdinalIgnoreCase))
                                      .ToArray();

        if (tiffFiles.Length == 0)
            throw new InvalidOperationException("No TIFF files found in the input directory.");

        // Prepare JPEG encoder with 90% quality
        ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                                                 .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);
        if (jpegCodec == null)
            throw new InvalidOperationException("JPEG codec not found.");

        EncoderParameters encoderParams = new EncoderParameters(1);
        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 90L);

        // Convert each TIFF to JPEG
        foreach (string tiffPath in tiffFiles)
        {
            using (Bitmap bitmap = new Bitmap(tiffPath))
            {
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(tiffPath);
                string jpegPath = Path.Combine(outputDir, fileNameWithoutExt + ".jpg");
                bitmap.Save(jpegPath, jpegCodec, encoderParams);
            }
        }

        // Validate output
        string[] jpegFiles = Directory.GetFiles(outputDir, "*.jpg", SearchOption.TopDirectoryOnly);
        if (jpegFiles.Length == 0)
            throw new InvalidOperationException("No JPEG files were created.");

        // Optionally, you could output the result paths (not required by the rules)
        // Console.WriteLine("Conversion completed. JPEG files created:");
        // foreach (var file in jpegFiles) Console.WriteLine(file);
    }

    private static void CreateSampleTiff(string filePath)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, width - 40, height - 40);
                }
            }
            bitmap.Save(filePath, ImageFormat.Tiff);
        }
    }
}
