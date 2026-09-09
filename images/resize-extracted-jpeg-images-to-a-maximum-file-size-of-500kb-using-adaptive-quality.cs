using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed file size for the resized JPEG images (500 KB).
    private const long MaxFileSizeBytes = 500 * 1024;

    public static void Main()
    {
        // Prepare deterministic folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputImagePath = Path.Combine(artifactsDir, "sample.jpg");

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int imgWidth = 800;
        const int imgHeight = 800;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill background with white and draw a simple rectangle.
            g.Clear(Aspose.Drawing.Color.White);
            using (Brush brush = new SolidBrush(Aspose.Drawing.Color.LightBlue))
            {
                g.FillRectangle(brush, 100, 100, 600, 600);
            }
            // Save as JPEG (default quality 95).
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Insert the sample image into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract JPEG images from the document and resize them adaptively.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Obtain original image bytes.
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Load the image into Aspose.Drawing.Bitmap.
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            using (Bitmap originalBitmap = new Bitmap(originalStream))
            {
                // Adaptive quality loop.
                int quality = 100;
                byte[] compressedBytes;
                do
                {
                    using (MemoryStream compressedStream = new MemoryStream())
                    {
                        // Set JPEG encoder with the current quality.
                        ImageCodecInfo jpegCodec = GetJpegCodec();
                        EncoderParameters encoderParams = new EncoderParameters(1);
                        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);
                        originalBitmap.Save(compressedStream, jpegCodec, encoderParams);
                        compressedBytes = compressedStream.ToArray();
                    }

                    // Reduce quality for the next iteration if needed.
                    if (compressedBytes.Length > MaxFileSizeBytes && quality > 10)
                        quality -= 10;
                    else
                        break;
                } while (true);

                // Save the resized image to a deterministic file name.
                string outputImagePath = Path.Combine(artifactsDir, $"extracted_{imageIndex}.jpg");
                File.WriteAllBytes(outputImagePath, compressedBytes);

                // Validation: ensure the file exists and meets size requirement.
                FileInfo info = new FileInfo(outputImagePath);
                if (!info.Exists)
                    throw new InvalidOperationException($"Failed to create output image: {outputImagePath}");
                if (info.Length > MaxFileSizeBytes)
                    throw new InvalidOperationException($"Image {outputImagePath} exceeds the maximum allowed size.");

                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 4. Indicate successful completion.
        // -----------------------------------------------------------------
        Console.WriteLine($"Processed {imageIndex} JPEG image(s). Resized images are saved in: {artifactsDir}");
    }

    // Helper method to retrieve the JPEG codec from Aspose.Drawing.Imaging.
    private static ImageCodecInfo GetJpegCodec()
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.FormatID == ImageFormat.Jpeg.Guid)
                return codec;
        }
        throw new InvalidOperationException("JPEG codec not found.");
    }
}
