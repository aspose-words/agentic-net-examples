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
        // Create a deterministic sample JPEG image.
        const string jpegPath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                g.DrawEllipse(new Pen(Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Create a Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Reload the document (optional, demonstrates loading workflow).
        Document loadedDoc = new Document(docPath);

        // Extract JPEG images and convert each to high‑resolution TIFF with LZW compression.
        int imageIndex = 0;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Load the JPEG image into a bitmap.
            using (MemoryStream ms = new MemoryStream(shape.ImageData.ImageBytes))
            {
                ms.Position = 0;
                using (Bitmap bmp = new Bitmap(ms))
                {
                    // Set high resolution (e.g., 300 DPI).
                    bmp.SetResolution(300f, 300f);

                    // Prepare TIFF encoder with LZW compression.
                    ImageCodecInfo tiffCodec = ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(c => c.FormatID == ImageFormat.Tiff.Guid);
                    if (tiffCodec == null)
                        throw new InvalidOperationException("TIFF encoder not found.");

                    EncoderParameters encoderParams = new EncoderParameters(1);
                    Encoder compressionEncoder = Encoder.Compression;
                    EncoderParameter compressionParam = new EncoderParameter(compressionEncoder, (long)EncoderValue.CompressionLZW);
                    encoderParams.Param[0] = compressionParam;

                    // Save the TIFF file.
                    string tiffPath = $"extracted-{imageIndex}.tiff";
                    bmp.Save(tiffPath, tiffCodec, encoderParams);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one TIFF file was created.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found and converted.");

        // Cleanup sample files (optional).
        // File.Delete(jpegPath);
        // File.Delete(docPath);
    }
}
