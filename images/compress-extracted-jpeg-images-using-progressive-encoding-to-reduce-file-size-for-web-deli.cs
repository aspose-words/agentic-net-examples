using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing.
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath);

        // 2. Build a Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleJpegPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // 3. Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        var jpegShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                  .Cast<Shape>()
                                  .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                                  .ToList();

        if (!jpegShapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        int index = 0;
        foreach (var shape in jpegShapes)
        {
            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Image image = Image.FromStream(originalStream))
                {
                    // Locate the JPEG codec.
                    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);
                    if (jpegCodec == null)
                        throw new InvalidOperationException("JPEG codec not found.");

                    // Set encoder parameters: quality = 70, progressive (interlaced) encoding.
                    using (EncoderParameters encoderParams = new EncoderParameters(2))
                    {
                        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 70L);
                        encoderParams.Param[1] = new EncoderParameter(Encoder.ScanMethod, (long)EncoderValue.ScanMethodInterlaced);

                        string compressedPath = Path.Combine(artifactsDir, $"compressed_{index}.jpg");
                        image.Save(compressedPath, jpegCodec, encoderParams);

                        // Validate that the compressed file exists and is smaller than the original.
                        FileInfo originalInfo = new FileInfo(sampleJpegPath);
                        FileInfo compressedInfo = new FileInfo(compressedPath);
                        if (!compressedInfo.Exists)
                            throw new InvalidOperationException($"Failed to create compressed image: {compressedPath}");
                        if (compressedInfo.Length >= originalInfo.Length)
                            Console.WriteLine($"Warning: compressed image {compressedPath} is not smaller than the original.");
                    }
                }
            }

            index++;
        }

        Console.WriteLine("Compression of extracted JPEG images completed successfully.");
    }

    // Helper method to create a deterministic sample JPEG image.
    private static void CreateSampleJpeg(string filePath)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
            {
                graphics.DrawEllipse(pen, 10, 10, width - 20, height - 20);
            }
            // Save with default quality (will be re‑compressed later).
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }
}
