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
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputImagePath = Path.Combine(artifactsDir, "sample.jpg");
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");

        // 1. Create a sample JPEG image.
        using (Bitmap bmp = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            g.DrawEllipse(new Pen(Aspose.Drawing.Color.DarkBlue, 5), 20, 20, 160, 160);
            bmp.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // 2. Insert the image into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // 3. Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                              .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        int index = 0;
        foreach (var shape in shapes)
        {
            // Save original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Prepare JPEG encoder with progressive (interlaced) encoding and quality.
                    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                                                             .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                    EncoderParameters encoderParams = new EncoderParameters(2);
                    // Quality = 70 (adjust as needed).
                    EncoderParameter qualityParam = new EncoderParameter(Encoder.Quality, 70L);
                    // Progressive (interlaced) scan method.
                    EncoderParameter scanParam = new EncoderParameter(Encoder.ScanMethod, (long)EncoderValue.ScanMethodInterlaced);
                    encoderParams.Param[0] = qualityParam;
                    encoderParams.Param[1] = scanParam;

                    // Save the compressed progressive JPEG.
                    string compressedPath = Path.Combine(artifactsDir, $"compressed_{index}.jpg");
                    originalBitmap.Save(compressedPath, jpegCodec, encoderParams);

                    // Validate that the compressed file exists and is smaller.
                    FileInfo originalInfo = new FileInfo(inputImagePath);
                    FileInfo compressedInfo = new FileInfo(compressedPath);
                    if (!compressedInfo.Exists)
                        throw new InvalidOperationException($"Compressed image not created: {compressedPath}");
                    if (compressedInfo.Length >= originalInfo.Length)
                        Console.WriteLine($"Warning: Compressed image {compressedInfo.Name} is not smaller than the original.");

                    index++;
                }
            }
        }

        // Indicate completion.
        Console.WriteLine("Image compression with progressive encoding completed.");
    }
}
