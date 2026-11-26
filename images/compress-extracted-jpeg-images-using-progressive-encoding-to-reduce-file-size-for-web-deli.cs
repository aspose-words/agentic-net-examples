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
        // Create a sample JPEG image.
        const int width = 200;
        const int height = 200;
        string sampleImagePath = "sample.jpg";

        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        graphics.DrawRectangle(new Pen(Color.Black), 10, 10, width - 20, height - 20);
        bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        graphics.Dispose();
        bitmap.Dispose();

        // Insert the image into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string originalDocPath = "original.docx";
        doc.Save(originalDocPath);

        // Load the document and compress each JPEG image using progressive encoding.
        Document loadedDoc = new Document(originalDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Export the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Bitmap imgBitmap = new Bitmap(originalStream))
                {
                    // Prepare progressive JPEG encoder parameters.
                    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);
                    if (jpegCodec == null)
                        throw new Exception("JPEG codec not found.");

                    EncoderParameters encoderParams = new EncoderParameters(1);
                    encoderParams.Param[0] = new EncoderParameter(Encoder.ScanMethod, (long)EncoderValue.ScanMethodInterlaced);

                    // Save the image with progressive encoding to a new stream.
                    using (MemoryStream progressiveStream = new MemoryStream())
                    {
                        imgBitmap.Save(progressiveStream, jpegCodec, encoderParams);
                        progressiveStream.Position = 0;

                        // Replace the shape's image data with the progressive JPEG.
                        shape.ImageData.SetImage(progressiveStream);

                        // Also write the progressive image to a file for verification.
                        string compressedImagePath = $"compressed-{imageIndex}.jpg";
                        File.WriteAllBytes(compressedImagePath, progressiveStream.ToArray());

                        // Validate that the file was created.
                        if (!File.Exists(compressedImagePath) || new FileInfo(compressedImagePath).Length == 0)
                            throw new Exception($"Failed to write progressive image '{compressedImagePath}'.");
                    }
                }
            }

            imageIndex++;
        }

        // Save the document with compressed images.
        string compressedDocPath = "Compressed.docx";
        loadedDoc.Save(compressedDocPath);

        // Final validation.
        if (!File.Exists(compressedDocPath) || imageIndex == 0)
            throw new Exception("Compression failed or no images were processed.");

        // Clean up temporary files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(originalDocPath);
    }
}
