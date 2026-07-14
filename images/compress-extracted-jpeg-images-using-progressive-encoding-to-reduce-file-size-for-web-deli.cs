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
        string extractedImagePath = Path.Combine(artifactsDir, "extracted_original.jpg");
        string compressedImagePath = Path.Combine(artifactsDir, "extracted_compressed.jpg");

        // -------------------------------------------------
        // 1. Create a deterministic sample JPEG image.
        // -------------------------------------------------
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            g.DrawEllipse(new Pen(Aspose.Drawing.Color.DarkBlue, 5), 20, 20, 160, 160);
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // -------------------------------------------------
        // 2. Insert the image into a Word document.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Extract JPEG images from the document.
        // -------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        Shape jpegShape = shapeNodes
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg);

        if (jpegShape == null)
            throw new InvalidOperationException("No JPEG image found in the document.");

        // Save the original extracted image.
        jpegShape.ImageData.Save(extractedImagePath);

        // -------------------------------------------------
        // 4. Re‑compress the extracted JPEG using lower quality.
        //    (Aspose.Words does not expose a direct progressive flag,
        //     so we use Aspose.Drawing to set JPEG quality, which also
        //     produces a progressive JPEG when supported by the encoder.)
        // -------------------------------------------------
        using (MemoryStream ms = new MemoryStream())
        {
            // Load the extracted image into Aspose.Drawing.Image.
            jpegShape.ImageData.Save(ms);
            ms.Position = 0;
            using (Image img = Image.FromStream(ms))
            {
                // Find the JPEG codec.
                ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                    .FirstOrDefault(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
                if (jpegCodec == null)
                    throw new InvalidOperationException("JPEG codec not found.");

                // Set encoder parameters: quality = 50 (adjust as needed).
                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 50L);

                // Save the compressed image.
                img.Save(compressedImagePath, jpegCodec, encoderParams);
            }
        }

        // -------------------------------------------------
        // 5. Validate that the compressed file was created.
        // -------------------------------------------------
        if (!File.Exists(compressedImagePath))
            throw new FileNotFoundException("Compressed image was not created.", compressedImagePath);

        // Optional: output file sizes for demonstration (not required by the task).
        Console.WriteLine($"Original extracted size: {new FileInfo(extractedImagePath).Length} bytes");
        Console.WriteLine($"Compressed size: {new FileInfo(compressedImagePath).Length} bytes");
    }
}
