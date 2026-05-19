using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
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
        string originalImagePath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bmp = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.White);
            g.DrawEllipse(new Pen(Color.Blue, 5), 20, 20, 160, 160);
            bmp.Save(originalImagePath, ImageFormat.Jpeg);
        }

        // 2. Insert the JPEG into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(originalImagePath);
        string docPath = Path.Combine(artifactsDir, "Original.docx");
        doc.Save(docPath);

        // 3. Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Image img = Image.FromStream(originalStream))
                {
                    // Prepare encoder parameters:
                    // - Quality = 50 (stronger compression)
                    // - ScanMethod = Interlaced (progressive JPEG)
                    EncoderParameters encoderParams = new EncoderParameters(2);
                    encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 50L);
                    encoderParams.Param[1] = new EncoderParameter(Encoder.ScanMethod, (long)EncoderValue.ScanMethodInterlaced);

                    // Find the JPEG codec.
                    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);
                    if (jpegCodec == null)
                        throw new InvalidOperationException("JPEG codec not found.");

                    // Save the compressed progressive JPEG.
                    string compressedPath = Path.Combine(artifactsDir, $"compressed_{imageIndex + 1}.jpg");
                    img.Save(compressedPath, jpegCodec, encoderParams);

                    // Validation: ensure the file exists and is smaller than the original.
                    FileInfo originalInfo = new FileInfo(originalImagePath);
                    FileInfo compressedInfo = new FileInfo(compressedPath);
                    if (!compressedInfo.Exists)
                        throw new FileNotFoundException("Compressed image was not created.", compressedPath);
                    if (compressedInfo.Length >= originalInfo.Length)
                        throw new InvalidOperationException("Compressed image is not smaller than the original.");

                    Console.WriteLine($"Image {imageIndex + 1} compressed: {originalInfo.Length} -> {compressedInfo.Length} bytes");
                }
            }

            imageIndex++;
        }

        // If no JPEG images were found, indicate it.
        if (imageIndex == 0)
            Console.WriteLine("No JPEG images were found in the document.");
    }
}
