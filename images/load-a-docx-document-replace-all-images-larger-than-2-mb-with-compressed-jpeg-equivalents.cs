using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ReplaceLargeImages
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Prepare a folder for all generated files.
        // -----------------------------------------------------------------
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Step 1: Create a large sample image (> 2 MB) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string largeImagePath = Path.Combine(artifactsDir, "large.bmp");
        int width = 4000;   // large enough to guarantee a big file size
        int height = 4000;

        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill with a solid colour – the exact colour is irrelevant.
            graphics.Clear(Aspose.Drawing.Color.LightGray);
            // Save as BMP (uncompressed) to ensure the file exceeds 2 MB.
            bitmap.Save(largeImagePath, ImageFormat.Bmp);
        }

        // Verify that the generated image is indeed larger than 2 MB.
        FileInfo largeInfo = new FileInfo(largeImagePath);
        if (largeInfo.Length <= 2 * 1024 * 1024)
            throw new Exception("Generated sample image is not larger than 2 MB.");

        // -----------------------------------------------------------------
        // Step 2: Create a DOCX document and insert the large image.
        // -----------------------------------------------------------------
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(largeImagePath);
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // Step 3: Load the document and replace images larger than 2 MB.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Obtain the current image bytes.
            byte[] originalBytes;
            using (MemoryStream tempStream = new MemoryStream())
            {
                shape.ImageData.Save(tempStream);
                originalBytes = tempStream.ToArray();
            }

            // Skip images that are already small enough.
            if (originalBytes.Length <= 2 * 1024 * 1024)
                continue;

            // Re‑encode the image as a JPEG with a moderate compression level.
            using (MemoryStream sourceStream = new MemoryStream(originalBytes))
            using (Bitmap bitmap = new Bitmap(sourceStream))
            using (MemoryStream jpegStream = new MemoryStream())
            {
                // Obtain the JPEG encoder.
                ImageCodecInfo jpegEncoder = ImageCodecInfo.GetImageEncoders()
                    .First(enc => enc.FormatID == ImageFormat.Jpeg.Guid);

                // Set compression quality (e.g., 50 %).
                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 50L);

                // Save the compressed JPEG to the memory stream.
                bitmap.Save(jpegStream, jpegEncoder, encoderParams);
                jpegStream.Position = 0; // Reset before feeding to Aspose.Words.

                // Replace the shape's image with the new JPEG data.
                shape.ImageData.SetImage(jpegStream);
            }
        }

        // -----------------------------------------------------------------
        // Step 4: Save the modified document.
        // -----------------------------------------------------------------
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");
        loadedDoc.Save(outputDocPath);

        // Simple validation.
        if (!File.Exists(outputDocPath))
            throw new Exception("Failed to create the output document.");

        Console.WriteLine($"Processing complete. Output saved to: {outputDocPath}");
    }
}
