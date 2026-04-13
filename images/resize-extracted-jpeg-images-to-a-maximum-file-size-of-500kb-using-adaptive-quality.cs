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
    // Maximum allowed file size for the resized JPEG (500 KB)
    private const long MaxFileSize = 500 * 1024;

    public static void Main()
    {
        // ---------------------------------------------------------------
        // 1. Create a deterministic sample JPEG image using Aspose.Drawing.
        // ---------------------------------------------------------------
        const string sampleImagePath = "input.jpg";
        const int sampleWidth = 800;
        const int sampleHeight = 800;

        // Create bitmap and fill with a solid color.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(sampleWidth, sampleHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.FromArgb(255, 70, 130, 180)); // SteelBlue background
        graphics.Dispose();

        // Save the bitmap as JPEG.
        bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        bitmap.Dispose();

        // ---------------------------------------------------------------
        // 2. Insert the sample image into a Word document using Aspose.Words.
        // ---------------------------------------------------------------
        const string docPath = "DocumentWithImage.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // InsertImage already appends the shape to the document, no extra AppendChild needed.
        Shape imageShape = builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 3. Load the document and extract each JPEG image, resizing it
        //    adaptively until its size does not exceed 500 KB.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading.

                // Load the image into Aspose.Drawing.Bitmap.
                using (Aspose.Drawing.Bitmap srcBitmap = new Aspose.Drawing.Bitmap(originalStream))
                {
                    // Prepare JPEG encoder.
                    Aspose.Drawing.Imaging.ImageCodecInfo jpegCodec = Aspose.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(c => c.FormatID == Aspose.Drawing.Imaging.ImageFormat.Jpeg.Guid);
                    if (jpegCodec == null)
                        throw new InvalidOperationException("JPEG codec not found.");

                    // Adaptive quality loop.
                    int quality = 100;
                    byte[] finalImageBytes = null;
                    while (quality >= 10)
                    {
                        using (MemoryStream resizedStream = new MemoryStream())
                        {
                            Aspose.Drawing.Imaging.EncoderParameters encoderParams = new Aspose.Drawing.Imaging.EncoderParameters(1);
                            encoderParams.Param[0] = new Aspose.Drawing.Imaging.EncoderParameter(Aspose.Drawing.Imaging.Encoder.Quality, (long)quality);

                            srcBitmap.Save(resizedStream, jpegCodec, encoderParams);
                            resizedStream.Position = 0; // Reset before size check.

                            if (resizedStream.Length <= MaxFileSize)
                            {
                                finalImageBytes = resizedStream.ToArray();
                                break;
                            }
                        }
                        quality -= 10; // Decrease quality and retry.
                    }

                    // If no quality satisfied the size constraint, use the lowest quality obtained.
                    if (finalImageBytes == null)
                    {
                        using (MemoryStream lowestStream = new MemoryStream())
                        {
                            Aspose.Drawing.Imaging.EncoderParameters encoderParams = new Aspose.Drawing.Imaging.EncoderParameters(1);
                            encoderParams.Param[0] = new Aspose.Drawing.Imaging.EncoderParameter(Aspose.Drawing.Imaging.Encoder.Quality, (long)10);
                            srcBitmap.Save(lowestStream, jpegCodec, encoderParams);
                            finalImageBytes = lowestStream.ToArray();
                        }
                    }

                    // Save the resized image to a deterministic file name.
                    string outputFileName = $"resized_image_{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                    File.WriteAllBytes(outputFileName, finalImageBytes);

                    // Validation: ensure the file exists and meets size requirement.
                    FileInfo info = new FileInfo(outputFileName);
                    if (!info.Exists)
                        throw new FileNotFoundException($"Failed to create {outputFileName}.");
                    if (info.Length > MaxFileSize)
                        throw new InvalidOperationException($"{outputFileName} exceeds the maximum allowed size.");

                    imageIndex++;
                }
            }
        }

        // If no images were processed, raise an exception as per validation rules.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were extracted and resized.");
    }
}
