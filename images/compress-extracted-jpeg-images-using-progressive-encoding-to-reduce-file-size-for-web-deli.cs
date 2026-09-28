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
        // Step 1: Create a deterministic sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath);

        // Step 2: Create a Word document and insert the sample JPEG.
        const string originalDocPath = "Original.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(originalDocPath);

        // Step 3: Load the document, extract JPEG images, recompress them using progressive encoding,
        // and replace the images in the document.
        Document loadedDoc = new Document(originalDocPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Only process JPEG images.
            string imageFormat = shape.ImageData.ImageType.ToString();
            if (imageFormat != "Jpeg") continue;

            // Extract original image bytes.
            byte[] originalBytes = shape.ImageData.ImageBytes;

            // Re‑encode the JPEG with progressive encoding and lower quality.
            byte[] compressedBytes = ReencodeJpegProgressive(originalBytes, quality: 50L);

            // Replace the image in the shape using a stream overload.
            using (MemoryStream ms = new MemoryStream(compressedBytes))
            {
                shape.ImageData.SetImage(ms);
            }

            // Save the compressed image to a file for verification.
            string extractedPath = $"extracted-{imageIndex}.jpg";
            File.WriteAllBytes(extractedPath, compressedBytes);
            imageIndex++;
        }

        // Validate that at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to compress.");

        // Step 4: Save the document with compressed images.
        const string compressedDocPath = "Compressed.docx";
        loadedDoc.Save(compressedDocPath);

        // Verify output files exist.
        if (!File.Exists(compressedDocPath))
            throw new FileNotFoundException("Compressed document was not created.", compressedDocPath);
        for (int i = 0; i < imageIndex; i++)
        {
            string path = $"extracted-{i}.jpg";
            if (!File.Exists(path))
                throw new FileNotFoundException("Compressed image file was not created.", path);
        }

        Console.WriteLine("Compression completed successfully.");
    }

    // Creates a simple 200x200 red JPEG image.
    private static void CreateSampleJpeg(string filePath)
    {
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 25, 25, 150, 150);
                }
            }

            // Save as JPEG (baseline) – will be recompressed later.
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Re‑encodes JPEG bytes using progressive encoding and the specified quality.
    private static byte[] ReencodeJpegProgressive(byte[] sourceBytes, long quality)
    {
        using (MemoryStream inputStream = new MemoryStream(sourceBytes))
        using (Bitmap bitmap = new Bitmap(inputStream))
        using (MemoryStream outputStream = new MemoryStream())
        {
            // Locate the JPEG encoder.
            ImageCodecInfo jpegEncoder = null;
            foreach (ImageCodecInfo codec in ImageCodecInfo.GetImageEncoders())
            {
                if (codec.FormatID == ImageFormat.Jpeg.Guid)
                {
                    jpegEncoder = codec;
                    break;
                }
            }

            if (jpegEncoder == null)
                throw new InvalidOperationException("JPEG encoder not found.");

            // Set encoder parameters: quality and progressive mode.
            EncoderParameters encoderParams = new EncoderParameters(2);
            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);
            encoderParams.Param[1] = new EncoderParameter(Encoder.RenderMethod, (long)EncoderValue.RenderProgressive);

            // Save the bitmap with the specified parameters.
            bitmap.Save(outputStream, jpegEncoder, encoderParams);
            return outputStream.ToArray();
        }
    }
}
