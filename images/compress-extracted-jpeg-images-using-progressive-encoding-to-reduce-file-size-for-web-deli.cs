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
        // Create a deterministic sample JPEG image.
        const int imgWidth = 200;
        const int imgHeight = 200;
        const string sampleImagePath = "sample.jpg";

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            using (Pen pen = new Pen(Color.Blue, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Insert the image into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "doc.docx";
        doc.Save(docPath);

        // Extract JPEG images, re‑encode them with progressive (interlaced) JPEG compression.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage || shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            using (MemoryStream imageStream = new MemoryStream())
            {
                // Save the original image bytes to a stream.
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image via Aspose.Drawing.
                using (Image img = Image.FromStream(imageStream))
                {
                    // Prepare encoder parameters for progressive JPEG.
                    Encoder scanEncoder = Encoder.ScanMethod;
                    EncoderParameter scanParam = new EncoderParameter(scanEncoder, (long)EncoderValue.ScanMethodInterlaced);

                    Encoder qualityEncoder = Encoder.Quality;
                    EncoderParameter qualityParam = new EncoderParameter(qualityEncoder, 75L); // Reasonable quality.

                    using (EncoderParameters encoderParams = new EncoderParameters(2))
                    {
                        encoderParams.Param[0] = scanParam;
                        encoderParams.Param[1] = qualityParam;

                        string outFileName = $"compressed_{extractedCount + 1}.jpg";
                        img.Save(outFileName, GetJpegCodec(), encoderParams);

                        // Validate that the file was created.
                        if (!File.Exists(outFileName))
                            throw new InvalidOperationException($"Failed to create compressed image '{outFileName}'.");

                        extractedCount++;
                    }
                }
            }
        }

        // Ensure at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No JPEG images were found to compress.");

        // Clean up temporary files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }

    // Retrieves the JPEG codec info required for Image.Save.
    private static ImageCodecInfo GetJpegCodec()
    {
        return ImageCodecInfo.GetImageEncoders()
            .First(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
    }
}
