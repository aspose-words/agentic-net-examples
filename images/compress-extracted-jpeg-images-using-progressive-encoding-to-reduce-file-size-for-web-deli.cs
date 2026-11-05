using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample JPEG image using Aspose.Drawing.
        const string sampleImagePath = "sample.jpg";
        using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.White);
                g.DrawRectangle(Aspose.Drawing.Pens.Black, 20, 20, 160, 160);
                g.DrawString("Test", new Aspose.Drawing.Font("Arial", 24), Aspose.Drawing.Brushes.Black, new Aspose.Drawing.PointF(50, 80));
            }
            bmp.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // Create a Word document and insert the JPEG image.
        const string inputDocPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // Reload the document for extraction.
        Document loadedDoc = new Document(inputDocPath);

        // Prepare output folder.
        const string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        // Iterate through all Shape nodes to find JPEG images.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Only process JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into Aspose.Drawing.Bitmap.
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    // Prepare encoder parameters for progressive JPEG with quality reduction.
                    EncoderParameters encoderParams = new EncoderParameters(1);
                    encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 75L);

                    // Get JPEG encoder.
                    ImageCodecInfo jpegEncoder = GetEncoder(Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                    if (jpegEncoder == null)
                        throw new InvalidOperationException("JPEG encoder not found.");

                    // Save compressed image to a new file.
                    string outputPath = Path.Combine(outputFolder, $"compressed-{imageIndex}.jpg");
                    using (MemoryStream compressedStream = new MemoryStream())
                    {
                        bitmap.Save(compressedStream, jpegEncoder, encoderParams);
                        File.WriteAllBytes(outputPath, compressedStream.ToArray());
                    }

                    // Validate that the file was created.
                    if (!File.Exists(outputPath))
                        throw new InvalidOperationException($"Failed to create compressed image: {outputPath}");
                }
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to compress.");

        // Clean up sample files.
        File.Delete(sampleImagePath);
        File.Delete(inputDocPath);
    }

    // Helper method to retrieve the encoder for a specific image format.
    private static ImageCodecInfo GetEncoder(Aspose.Drawing.Imaging.ImageFormat format)
    {
        ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
        foreach (ImageCodecInfo codec in codecs)
        {
            if (codec.FormatID == format.Guid)
                return codec;
        }
        return null;
    }
}
