using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image.
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(jpegPath);

        // 2. Create a Word document that contains the JPEG image.
        string docPath = Path.Combine(artifactsDir, "DocumentWithJpeg.docx");
        CreateDocumentWithImage(jpegPath, docPath);

        // 3. Load the document and extract JPEG images.
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int jpegCount = 0;
        int extractedIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            if (shape.ImageData.ImageType == ImageType.Jpeg)
            {
                jpegCount++;

                // a. Save the extracted JPEG to a temporary file.
                string extractedJpeg = Path.Combine(artifactsDir, $"extracted_{extractedIndex}.jpg");
                shape.ImageData.Save(extractedJpeg);
                if (!File.Exists(extractedJpeg))
                    throw new InvalidOperationException($"Failed to save extracted JPEG: {extractedJpeg}");

                // b. Convert the JPEG to high‑quality WebP.
                string webpPath = Path.Combine(artifactsDir, $"converted_{extractedIndex}.webp");
                ConvertJpegToWebp(extractedJpeg, webpPath);
                if (!File.Exists(webpPath))
                    throw new InvalidOperationException($"Failed to create WebP file: {webpPath}");

                extractedIndex++;
            }
        }

        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were found in the document.");

        Console.WriteLine($"Processed {jpegCount} JPEG image(s). WebP files are saved in: {artifactsDir}");
    }

    // Creates a deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (Brush brush = new SolidBrush(Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, width - 40, height - 40);
                }
            }

            // Save as JPEG with high quality.
            ImageCodecInfo jpegCodec = GetEncoder(ImageFormat.Jpeg);
            EncoderParameters encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 100L);
            bitmap.Save(filePath, jpegCodec, encoderParams);
        }
    }

    // Inserts the given image into a new Word document.
    private static void CreateDocumentWithImage(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Converts a JPEG file to WebP using Aspose.Words rendering pipeline.
    private static void ConvertJpegToWebp(string jpegPath, string webpPath)
    {
        // Load the JPEG into a temporary document.
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.InsertImage(jpegPath);

        // Configure ImageSaveOptions for WebP with high quality.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.WebP);
        // The JpegQuality property does not affect WebP, but we can set the compression level via ImageQuality if needed.
        // Here we rely on default high quality.

        // Save the rendered page as a WebP image.
        tempDoc.Save(webpPath, options);
    }

    // Helper to obtain the JPEG encoder.
    private static ImageCodecInfo GetEncoder(ImageFormat format)
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.FormatID == format.Guid)
                return codec;
        }
        throw new InvalidOperationException("JPEG encoder not found.");
    }
}
