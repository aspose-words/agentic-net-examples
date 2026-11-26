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
        // -----------------------------------------------------------------
        // 1. Create a deterministic BMP image that will be inserted into the document.
        // -----------------------------------------------------------------
        const string bmpPath = "sample.bmp";
        const int bmpWidth = 800;
        const int bmpHeight = 600;

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(bmpWidth, bmpHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 50, 50, bmpWidth - 100, bmpHeight - 100);
                }
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
                {
                    g.DrawString("Sample BMP", font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red),
                        new Aspose.Drawing.PointF(150, 250));
                }
            }
            // Save the BMP to disk – required before inserting into the document.
            bitmap.Save(bmpPath);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the BMP image.
        // -----------------------------------------------------------------
        const string docPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Reload the document to simulate a fresh load.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // 4. Extract each image, resize it (quality scaling) so the file size
        //    does not exceed 200 KB, and save the result as JPEG.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        const int maxFileSize = 200 * 1024; // 200 KB

        // JPEG encoder used for quality scaling.
        Aspose.Drawing.Imaging.ImageCodecInfo jpegEncoder = Aspose.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
            .First(enc => enc.FormatID == Aspose.Drawing.Imaging.ImageFormat.Jpeg.Guid);

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue; // Skip shapes without an image.

            // Retrieve the original image bytes from the document.
            byte[] originalBytes = shape.ImageData.ImageBytes;

            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            {
                using (Aspose.Drawing.Bitmap originalBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    byte[] resizedBytes = null;

                    // Try decreasing JPEG quality until the size requirement is met.
                    for (int quality = 100; quality >= 10; quality -= 10)
                    {
                        using (MemoryStream resizedStream = new MemoryStream())
                        {
                            Aspose.Drawing.Imaging.EncoderParameters encoderParams = new Aspose.Drawing.Imaging.EncoderParameters(1);
                            encoderParams.Param[0] = new Aspose.Drawing.Imaging.EncoderParameter(
                                Aspose.Drawing.Imaging.Encoder.Quality, (long)quality);

                            originalBitmap.Save(resizedStream, jpegEncoder, encoderParams);

                            if (resizedStream.Length <= maxFileSize)
                            {
                                resizedBytes = resizedStream.ToArray();
                                break;
                            }
                        }
                    }

                    // If even the lowest quality is too large, fall back to the lowest quality anyway.
                    if (resizedBytes == null)
                    {
                        using (MemoryStream finalStream = new MemoryStream())
                        {
                            Aspose.Drawing.Imaging.EncoderParameters encoderParams = new Aspose.Drawing.Imaging.EncoderParameters(1);
                            encoderParams.Param[0] = new Aspose.Drawing.Imaging.EncoderParameter(
                                Aspose.Drawing.Imaging.Encoder.Quality, 10L);
                            originalBitmap.Save(finalStream, jpegEncoder, encoderParams);
                            resizedBytes = finalStream.ToArray();
                        }
                    }

                    // Save the resized JPEG to a deterministic file name.
                    string outputPath = $"resized-{imageIndex}.jpg";
                    File.WriteAllBytes(outputPath, resizedBytes);

                    // Validate that the file was written and is not empty.
                    if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
                        throw new InvalidOperationException($"Failed to write resized image '{outputPath}'.");

                    imageIndex++;
                }
            }
        }

        // Ensure that at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found to process.");
    }
}
