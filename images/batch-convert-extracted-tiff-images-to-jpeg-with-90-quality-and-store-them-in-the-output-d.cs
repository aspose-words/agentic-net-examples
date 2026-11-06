using System;
using System.IO;
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

        // Create sample TIFF images
        for (int i = 1; i <= 3; i++)
        {
            string tiffPath = Path.Combine(inputFolder, $"sample{i}.tif");
            using (Bitmap bmp = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    g.FillRectangle(new SolidBrush(Aspose.Drawing.Color.FromArgb(50 * i, 100, 150)), 20, 20, 160, 160);
                }
                bmp.Save(tiffPath, ImageFormat.Tiff);
            }
        }

        // Batch convert TIFF to JPEG with 90% quality
        string[] tiffFiles = Directory.GetFiles(inputFolder, "*.tif");
        foreach (string tiffFile in tiffFiles)
        {
            using (Bitmap bitmap = new Bitmap(tiffFile))
            {
                string jpegFileName = Path.GetFileNameWithoutExtension(tiffFile) + ".jpg";
                string jpegPath = Path.Combine(outputFolder, jpegFileName);

                // Set JPEG quality to 90%
                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 90L);

                ImageCodecInfo jpegCodec = GetEncoder(ImageFormat.Jpeg);
                if (jpegCodec == null)
                    throw new InvalidOperationException("JPEG encoder not found.");

                bitmap.Save(jpegPath, jpegCodec, encoderParams);
            }
        }

        // Validate output
        int jpegCount = Directory.GetFiles(outputFolder, "*.jpg").Length;
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG files were created.");

        Console.WriteLine($"Conversion completed. {jpegCount} JPEG file(s) saved to '{outputFolder}'.");
    }

    private static ImageCodecInfo GetEncoder(ImageFormat format)
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.FormatID == format.Guid)
                return codec;
        }
        return null;
    }
}
