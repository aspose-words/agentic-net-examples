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
        // Define file and folder paths
        string inputImagePath = "input.png";
        string docPath = "sample.docx";
        string extractedFolder = "extracted";
        string compressedFolder = "compressed";

        // Ensure output directories exist
        Directory.CreateDirectory(extractedFolder);
        Directory.CreateDirectory(compressedFolder);

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the image multiple times
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(inputImagePath);
            builder.Writeln(); // Add a line break between images
        }
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        var extractedFiles = new System.Collections.Generic.List<string>();

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedPath = Path.Combine(extractedFolder, $"extracted-{imageIndex}.png");
                shape.ImageData.Save(extractedPath);
                extractedFiles.Add(extractedPath);
                imageIndex++;
            }
        }

        if (extractedFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Apply lossless PNG compression to each extracted image
        // -------------------------------------------------
        var pngEncoder = ImageCodecInfo.GetImageEncoders()
            .FirstOrDefault(enc => enc.FormatID == ImageFormat.Png.Guid);
        if (pngEncoder == null)
            throw new InvalidOperationException("PNG encoder not found.");

        var compressionEncoder = Encoder.Compression;
        var compressionLevel = (long)9; // Maximum lossless compression

        var encoderParams = new EncoderParameters(1);
        encoderParams.Param[0] = new EncoderParameter(compressionEncoder, compressionLevel);

        int compressedIndex = 0;
        foreach (string extractedPath in extractedFiles)
        {
            using (Bitmap bmp = new Bitmap(extractedPath))
            {
                string compressedPath = Path.Combine(compressedFolder, $"compressed-{compressedIndex}.png");
                bmp.Save(compressedPath, pngEncoder, encoderParams);

                // -------------------------------------------------
                // 5. Compare file sizes and output statistics
                // -------------------------------------------------
                long originalSize = new FileInfo(extractedPath).Length;
                long compressedSize = new FileInfo(compressedPath).Length;
                double reductionPercent = originalSize == 0 ? 0 :
                    100.0 * (originalSize - compressedSize) / originalSize;

                Console.WriteLine($"Image {compressedIndex}: Original = {originalSize} bytes, " +
                                  $"Compressed = {compressedSize} bytes, " +
                                  $"Reduction = {reductionPercent:F2}%");
                compressedIndex++;
            }
        }

        // Validate that compressed files were created
        if (compressedIndex == 0)
            throw new InvalidOperationException("No images were compressed.");
    }
}
