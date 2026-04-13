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
        // Create a deterministic BMP image larger than 200 KB.
        const string bmpPath = "input.bmp";
        const int bmpWidth = 800;
        const int bmpHeight = 800;
        using (Bitmap bmp = new Bitmap(bmpWidth, bmpHeight))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.White);
            // Draw a simple pattern to increase file size.
            for (int i = 0; i < 100; i++)
            {
                g.FillRectangle(new SolidBrush(Color.FromArgb(i * 2 % 256, i * 5 % 256, i * 3 % 256)),
                                 i * 5, i * 5, bmpWidth - i * 10, bmpHeight - i * 10);
            }
            bmp.Save(bmpPath);
        }

        // Insert the BMP into a Word document.
        const string docPath = "DocumentWithImage.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        doc.Save(docPath);

        // Load the document and process each image.
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .Where(s => s.HasImage)
                              .ToList();

        if (!shapes.Any())
            throw new Exception("No images found in the document.");

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reuse.

                // Load the image with Aspose.Drawing.
                using (Image img = Image.FromStream(originalStream))
                {
                    // Determine if the image already satisfies the size constraint.
                    const long maxSizeBytes = 200 * 1024;
                    byte[] jpegBytes;
                    int quality = 100;

                    // Loop decreasing JPEG quality until size requirement is met.
                    do
                    {
                        using (MemoryStream jpegStream = new MemoryStream())
                        {
                            // Obtain JPEG encoder.
                            ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                                .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);
                            if (jpegCodec == null)
                                throw new Exception("JPEG codec not found.");

                            // Set encoder parameters for quality.
                            EncoderParameters encoderParams = new EncoderParameters(1);
                            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);

                            // Save with current quality.
                            img.Save(jpegStream, jpegCodec, encoderParams);
                            jpegBytes = jpegStream.ToArray();
                        }

                        if (jpegBytes.Length <= maxSizeBytes || quality <= 10)
                            break;

                        quality -= 10; // Reduce quality and retry.
                    } while (true);

                    // Save the resized image to a deterministic file.
                    string resizedPath = $"resized_image_{imageIndex}.jpg";
                    File.WriteAllBytes(resizedPath, jpegBytes);

                    // Validation.
                    if (!File.Exists(resizedPath) || new FileInfo(resizedPath).Length == 0)
                        throw new Exception($"Failed to save resized image '{resizedPath}'.");

                    imageIndex++;
                }
            }
        }

        // Ensure at least one resized image was produced.
        if (imageIndex == 0)
            throw new Exception("No resized images were created.");
    }
}
