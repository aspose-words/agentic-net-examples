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
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputImages");
        string outputFolder = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample BMP images
        for (int i = 1; i <= 3; i++)
        {
            string bmpPath = Path.Combine(inputFolder, $"image{i}.bmp");
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            g.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 2))
            {
                g.DrawRectangle(pen, 10, 10, 80, 80);
            }
            bitmap.Save(bmpPath, Aspose.Drawing.Imaging.ImageFormat.Bmp);
            g.Dispose();
            bitmap.Dispose();
        }

        // Batch convert BMP to JPEG with 80% quality
        string[] bmpFiles = Directory.GetFiles(inputFolder, "*.bmp");
        foreach (string bmpFile in bmpFiles)
        {
            // Load BMP into a temporary Word document to obtain Shape.ImageData
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape shape = builder.InsertImage(bmpFile);
            if (!shape.HasImage)
                throw new Exception($"Shape does not contain an image for file {bmpFile}");

            // Save image data to a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // reset before reading

                // Load bitmap from stream
                using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(ms))
                {
                    // Prepare JPEG encoder with 80% quality
                    Aspose.Drawing.Imaging.ImageCodecInfo jpegCodec = GetEncoder(Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                    if (jpegCodec == null)
                        throw new Exception("JPEG encoder not found.");

                    Aspose.Drawing.Imaging.EncoderParameters encoderParams = new Aspose.Drawing.Imaging.EncoderParameters(1);
                    encoderParams.Param[0] = new Aspose.Drawing.Imaging.EncoderParameter(Aspose.Drawing.Imaging.Encoder.Quality, 80L);

                    // Save JPEG
                    string jpegPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(bmpFile) + ".jpg");
                    bmp.Save(jpegPath, jpegCodec, encoderParams);

                    Console.WriteLine($"Converted '{Path.GetFileName(bmpFile)}' to '{Path.GetFileName(jpegPath)}'.");
                }
            }
        }

        // Validate that at least one JPEG was created
        string[] jpegFiles = Directory.GetFiles(outputFolder, "*.jpg");
        if (jpegFiles.Length == 0)
            throw new Exception("No JPEG files were created during conversion.");
    }

    // Helper to get the appropriate image encoder
    private static Aspose.Drawing.Imaging.ImageCodecInfo GetEncoder(Aspose.Drawing.Imaging.ImageFormat format)
    {
        Aspose.Drawing.Imaging.ImageCodecInfo[] codecs = Aspose.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();
        foreach (Aspose.Drawing.Imaging.ImageCodecInfo codec in codecs)
        {
            if (codec.FormatID == format.Guid)
                return codec;
        }
        return null;
    }
}
