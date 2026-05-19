using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Maximum allowed file size for the resized JPEG (500 KB).
    private const long MaxFileSizeBytes = 500 * 1024;

    public static void Main()
    {
        // 1. Create a deterministic sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath, 800, 800);

        // 2. Insert the sample image into a Word document.
        const string docPath = "document.docx";
        InsertImageIntoDocument(sampleImagePath, docPath);

        // 3. Load the document and extract all JPEG images, resizing them adaptively.
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Only process JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract original image bytes.
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Resize adaptively to meet the size constraint.
            byte[] resizedBytes = ResizeJpegToTargetSize(originalBytes, MaxFileSizeBytes);

            // Save the resized image to a deterministic file name.
            string outputFileName = $"extracted_{imageIndex}.jpg";
            File.WriteAllBytes(outputFileName, resizedBytes);

            // Validation: ensure the file exists and respects the size limit.
            FileInfo info = new FileInfo(outputFileName);
            if (!info.Exists)
                throw new InvalidOperationException($"Failed to create {outputFileName}.");
            if (info.Length > MaxFileSizeBytes)
                throw new InvalidOperationException($"{outputFileName} exceeds the maximum allowed size.");

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to process.");
    }

    // Creates a solid‑color JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            // Save with default quality (100) to ensure a reasonably large file.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Inserts an image file into a new Word document and saves it.
    private static void InsertImageIntoDocument(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Resizes a JPEG image by lowering its quality until it fits within the target size.
    private static byte[] ResizeJpegToTargetSize(byte[] imageBytes, long maxSizeBytes)
    {
        // Load the image from the original byte array.
        using (MemoryStream inputStream = new MemoryStream(imageBytes))
        using (Bitmap bitmap = new Bitmap(inputStream))
        {
            // Start with the highest quality.
            long quality = 100;
            const long minQuality = 10;

            while (true)
            {
                using (MemoryStream outputStream = new MemoryStream())
                {
                    // Set JPEG encoder parameters for the current quality.
                    ImageCodecInfo jpegCodec = GetJpegCodec();
                    EncoderParameters encoderParams = new EncoderParameters(1);
                    encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);

                    // Save the bitmap to the stream using the specified quality.
                    bitmap.Save(outputStream, jpegCodec, encoderParams);
                    byte[] result = outputStream.ToArray();

                    // If the size is acceptable or we have reached the minimum quality, return.
                    if (result.Length <= maxSizeBytes || quality <= minQuality)
                        return result;

                    // Reduce quality and try again.
                    quality -= 10;
                    if (quality < minQuality)
                        quality = minQuality;
                }
            }
        }
    }

    // Retrieves the JPEG codec info required for encoding with quality settings.
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
