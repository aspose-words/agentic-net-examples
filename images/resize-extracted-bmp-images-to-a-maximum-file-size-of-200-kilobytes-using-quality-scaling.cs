using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed file size in bytes (200 KB)
    private const long MaxFileSize = 200 * 1024;

    public static void Main()
    {
        // Create a sample BMP image.
        const string inputBmpPath = "sample.bmp";
        CreateSampleBmp(inputBmpPath, 800, 600); // 800x600 pixels

        // Create a Word document and insert the BMP image.
        const string docPath = "DocumentWithBmp.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputBmpPath);
        doc.Save(docPath);

        // Load the document and process each BMP image.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Process only BMP images.
            if (shape.ImageData.ImageType != ImageType.Bmp) continue;

            // Obtain the original image bytes.
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // If the image already satisfies the size requirement, just save it.
            if (originalBytes.Length <= MaxFileSize)
            {
                string outPath = $"ExtractedImage_{imageIndex}.bmp";
                shape.ImageData.Save(outPath);
                ValidateFile(outPath);
                imageIndex++;
                continue;
            }

            // Load the BMP into Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(originalBytes))
            using (Bitmap bitmap = new Bitmap(ms))
            {
                // Encode the bitmap as JPEG with a quality setting.
                // Start with high quality and decrease until the size constraint is met.
                int quality = 100;
                byte[] jpegBytes;

                do
                {
                    using (MemoryStream jpegStream = new MemoryStream())
                    {
                        // Set JPEG encoder parameters.
                        EncoderParameters encoderParams = new EncoderParameters(1);
                        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);
                        ImageCodecInfo jpegCodec = GetEncoder(ImageFormat.Jpeg);
                        bitmap.Save(jpegStream, jpegCodec, encoderParams);
                        jpegBytes = jpegStream.ToArray();
                    }

                    // Reduce quality stepwise if still too large.
                    quality -= 10;
                }
                while (jpegBytes.Length > MaxFileSize && quality > 0);

                // Save the resulting image.
                string outputPath = $"ResizedImage_{imageIndex}.jpg";
                File.WriteAllBytes(outputPath, jpegBytes);
                ValidateFile(outputPath);
                imageIndex++;
            }
        }

        Console.WriteLine("Processing completed.");
    }

    // Creates a deterministic BMP file with a solid background.
    private static void CreateSampleBmp(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            bitmap.Save(filePath, ImageFormat.Bmp);
        }
    }

    // Retrieves the encoder for a specific image format.
    private static ImageCodecInfo GetEncoder(ImageFormat format)
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.FormatID == format.Guid)
                return codec;
        }
        throw new InvalidOperationException("Encoder not found for the specified format.");
    }

    // Validates that a file exists and is non‑empty.
    private static void ValidateFile(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length == 0)
            throw new InvalidOperationException($"Failed to create output file: {path}");
    }
}
