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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                g.DrawEllipse(new Pen(Aspose.Drawing.Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bmp.Save(jpegPath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the JPEG image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract JPEG images from the document.
        // -----------------------------------------------------------------
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .Cast<Shape>()
                        .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                        .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        int imageIndex = 0;
        foreach (var shape in shapes)
        {
            // Save the original JPEG to a memory stream.
            using (MemoryStream jpegStream = new MemoryStream())
            {
                shape.ImageData.Save(jpegStream);
                jpegStream.Position = 0;

                // Load the JPEG into a Bitmap.
                using (Bitmap bitmap = new Bitmap(jpegStream))
                {
                    // Set a high resolution (e.g., 300 DPI) for archival quality.
                    bitmap.SetResolution(300f, 300f);

                    // Prepare TIFF encoder with LZW compression.
                    ImageCodecInfo tiffCodec = GetEncoder(ImageFormat.Tiff);
                    EncoderParameters encoderParams = new EncoderParameters(1);
                    EncoderParameter compressionParam = new EncoderParameter(
                        Encoder.Compression,
                        (long)EncoderValue.CompressionLZW);
                    encoderParams.Param[0] = compressionParam;

                    // Save as TIFF with LZW compression.
                    string tiffPath = Path.Combine(artifactsDir, $"ExtractedImage_{imageIndex}.tiff");
                    bitmap.Save(tiffPath, tiffCodec, encoderParams);

                    // Validate that the TIFF file was created.
                    if (!File.Exists(tiffPath))
                        throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
                }
            }
            imageIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Demonstrate saving a whole document page as high‑resolution TIFF
        //    with LZW compression (optional extra step).
        // -----------------------------------------------------------------
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw,
            Resolution = 300 // 300 DPI
        };
        string docTiffPath = Path.Combine(artifactsDir, "DocumentPage.tiff");
        doc.Save(docTiffPath, tiffOptions);
    }

    // Helper method to obtain the encoder for a specific image format.
    private static ImageCodecInfo GetEncoder(ImageFormat format)
    {
        return ImageCodecInfo.GetImageEncoders()
                             .FirstOrDefault(codec => codec.FormatID == format.Guid)
               ?? throw new InvalidOperationException($"Encoder not found for format {format}.");
    }
}
