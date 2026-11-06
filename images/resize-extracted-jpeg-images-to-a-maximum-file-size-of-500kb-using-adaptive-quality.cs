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
        // ------------------------------------------------------------
        // 1. Create a deterministic sample JPEG image using Aspose.Drawing.
        // ------------------------------------------------------------
        const string sampleImagePath = "sample.jpg";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(800, 600);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
        {
            graphics.DrawString(
                "Sample",
                font,
                new SolidBrush(Aspose.Drawing.Color.Black),
                new PointF(10, 10));
        }
        graphics.Dispose();
        bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        bitmap.Dispose();

        // ------------------------------------------------------------
        // 2. Insert the JPEG image into a new Word document.
        // ------------------------------------------------------------
        const string inputDocPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // ------------------------------------------------------------
        // 3. Load the document for processing.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // ------------------------------------------------------------
        // 4. Prepare output folder for resized images.
        // ------------------------------------------------------------
        const string outputFolder = "ResizedImages";
        Directory.CreateDirectory(outputFolder);

        // ------------------------------------------------------------
        // 5. Iterate over all Shape nodes that contain images.
        // ------------------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        const long maxFileSize = 500 * 1024; // 500 KB

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Retrieve the original image bytes.
            byte[] originalBytes = shape.ImageData.ImageBytes;

            // Simple JPEG detection (first two bytes 0xFF,0xD8).
            if (originalBytes.Length < 2 || originalBytes[0] != 0xFF || originalBytes[1] != 0xD8)
                continue; // Not a JPEG, skip.

            // Load the image using Aspose.Drawing.
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            {
                originalStream.Position = 0;
                using (Image image = Image.FromStream(originalStream))
                {
                    // Adaptive quality loop: try decreasing quality until size fits.
                    byte[] bestBytes = null;
                    for (int quality = 90; quality >= 10; quality -= 10)
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Jpeg.Guid);

                            EncoderParameters encoderParams = new EncoderParameters(1);
                            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)quality);

                            image.Save(ms, jpegCodec, encoderParams);
                            ms.Position = 0;
                            byte[] candidate = ms.ToArray();

                            if (candidate.Length <= maxFileSize)
                            {
                                bestBytes = candidate;
                                break; // Acceptable size found.
                            }
                        }
                    }

                    // If no acceptable size was found, use the smallest quality attempted.
                    if (bestBytes == null)
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                                .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                            EncoderParameters encoderParams = new EncoderParameters(1);
                            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)10);
                            image.Save(ms, jpegCodec, encoderParams);
                            bestBytes = ms.ToArray();
                        }
                    }

                    // Save the resized image to the output folder.
                    string outputPath = Path.Combine(outputFolder, $"resized-{imageIndex}.jpg");
                    File.WriteAllBytes(outputPath, bestBytes);

                    // Validate that the file was created and respects the size limit.
                    if (!File.Exists(outputPath) || new FileInfo(outputPath).Length > maxFileSize)
                        throw new InvalidOperationException($"Failed to create resized image within size limit: {outputPath}");
                }
            }

            imageIndex++;
        }

        // ------------------------------------------------------------
        // 6. Validation: ensure at least one image was processed.
        // ------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found and processed in the document.");
    }
}
