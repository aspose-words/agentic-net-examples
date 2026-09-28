using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using ImageCodecInfo = Aspose.Drawing.Imaging.ImageCodecInfo;
using Encoder = Aspose.Drawing.Imaging.Encoder;
using EncoderParameter = Aspose.Drawing.Imaging.EncoderParameter;
using EncoderParameters = Aspose.Drawing.Imaging.EncoderParameters;

public class Program
{
    // Maximum allowed file size in bytes (500 KB)
    private const long MaxFileSize = 500 * 1024;

    public static void Main()
    {
        // Step 1: Create a deterministic sample JPEG image.
        const string sampleImagePath = "input.jpg";
        CreateSampleJpeg(sampleImagePath, 800, 600);

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        CreateDocumentWithImage(docPath, sampleImagePath);

        // Step 3: Load the document and extract JPEG images.
        Document doc = new Document(docPath);
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage) continue;

            // Extract the image to a memory stream.
            using (MemoryStream extractedStream = new MemoryStream())
            {
                shape.ImageData.Save(extractedStream);
                extractedStream.Position = 0;

                // Load the extracted image into Aspose.Drawing.Image.
                using (Image originalImage = Image.FromStream(extractedStream))
                {
                    // Adaptive quality compression to meet the size limit.
                    using (Image resizedImage = AdaptiveResizeJpeg(originalImage, MaxFileSize))
                    {
                        // Save the resized image to a deterministic file name.
                        string outputPath = $"extracted_resized_{imageIndex}.jpg";
                        resizedImage.Save(outputPath, GetJpegCodecInfo(), GetEncoderParameters(100L));

                        // Validate the output file.
                        FileInfo info = new FileInfo(outputPath);
                        if (!info.Exists || info.Length == 0)
                            throw new InvalidOperationException($"Failed to create output image: {outputPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Creates a simple JPEG image with white background.
    private static void CreateSampleJpeg(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual content.
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 50, 50, width - 100, height - 100);
                }
            }

            // Save as JPEG with high quality.
            bitmap.Save(path, GetJpegCodecInfo(), GetEncoderParameters(100L));
        }

        // Validate the image file exists.
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample image at {path}");
    }

    // Creates a Word document and inserts the specified image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create document at {docPath}");
    }

    // Returns the JPEG codec info.
    private static ImageCodecInfo GetJpegCodecInfo()
    {
        return ImageCodecInfo.GetImageEncoders()
            .FirstOrDefault(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
    }

    // Returns encoder parameters with the specified quality.
    private static EncoderParameters GetEncoderParameters(long quality)
    {
        EncoderParameters parameters = new EncoderParameters(1);
        parameters.Param[0] = new EncoderParameter(Encoder.Quality, quality);
        return parameters;
    }

    // Performs adaptive quality reduction to fit within maxSize bytes.
    private static Image AdaptiveResizeJpeg(Image original, long maxSize)
    {
        // Start with high quality and decrement until size constraint is met.
        for (int quality = 100; quality >= 10; quality -= 10)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                original.Save(ms, GetJpegCodecInfo(), GetEncoderParameters(quality));
                if (ms.Length <= maxSize)
                {
                    ms.Position = 0;
                    return Image.FromStream(ms);
                }
            }
        }

        // If all attempts exceed the limit, return the lowest quality version.
        using (MemoryStream finalMs = new MemoryStream())
        {
            original.Save(finalMs, GetJpegCodecInfo(), GetEncoderParameters(10L));
            finalMs.Position = 0;
            return Image.FromStream(finalMs);
        }
    }
}
