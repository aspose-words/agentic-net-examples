using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed file size in bytes (500 KB)
    private const long MaxFileSize = 500 * 1024;

    public static void Main()
    {
        // Prepare directories
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputImagePath = Path.Combine(artifactsDir, "sample.jpg");
        string docPath = Path.Combine(artifactsDir, "document.docx");

        // 1. Create a sample JPEG image using Aspose.Drawing
        CreateSampleJpeg(inputImagePath, 800, 800);

        // 2. Insert the image into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // 3. Load the document and extract JPEG images
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                              .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        int imageIndex = 0;
        foreach (var shape in shapes)
        {
            // Extract original image bytes
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Resize adaptively to meet the size constraint
            byte[] resizedBytes = ResizeJpegAdaptive(originalBytes, MaxFileSize);

            // Save the resized image
            string outputPath = Path.Combine(artifactsDir, $"resized_{imageIndex}.jpg");
            File.WriteAllBytes(outputPath, resizedBytes);

            // Validate output
            FileInfo info = new FileInfo(outputPath);
            if (info.Length > MaxFileSize)
                throw new InvalidOperationException($"Resized image {outputPath} exceeds the maximum allowed size.");

            imageIndex++;
        }

        // All done – the resized images are stored in the Artifacts folder.
    }

    // Creates a deterministic JPEG image with a solid color background.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.FromArgb(255, 70, 130, 180)); // SteelBlue background
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Performs adaptive quality reduction to fit the image within the target size.
    private static byte[] ResizeJpegAdaptive(byte[] sourceBytes, long maxSize)
    {
        // Load the source image into Aspose.Drawing.Image
        using (MemoryStream sourceStream = new MemoryStream(sourceBytes))
        using (Image image = Image.FromStream(sourceStream))
        {
            // Find the JPEG encoder
            ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                                                     .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);
            if (jpegCodec == null)
                throw new InvalidOperationException("JPEG encoder not found.");

            // Start with high quality and decrease until size constraint is met
            for (int quality = 100; quality >= 10; quality -= 10)
            {
                using (EncoderParameters encoderParams = new EncoderParameters(1))
                using (MemoryStream outputStream = new MemoryStream())
                {
                    encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);
                    image.Save(outputStream, jpegCodec, encoderParams);
                    if (outputStream.Length <= maxSize)
                        return outputStream.ToArray();
                }
            }

            // If none of the quality levels satisfy the constraint, return the lowest quality version
            using (EncoderParameters encoderParams = new EncoderParameters(1))
            using (MemoryStream outputStream = new MemoryStream())
            {
                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 10L);
                image.Save(outputStream, jpegCodec, encoderParams);
                return outputStream.ToArray();
            }
        }
    }
}
