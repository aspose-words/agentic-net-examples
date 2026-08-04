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
        // Directories for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            graphics.DrawEllipse(new Pen(Color.DarkBlue, 5), 20, 20, 160, 160);
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // 2. Create a Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "Document.docx");
        doc.Save(docPath);

        // 3. Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                              .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
                              .ToList();

        if (shapes.Count == 0)
            throw new InvalidOperationException("No JPEG images were found in the document.");

        int index = 0;
        foreach (var shape in shapes)
        {
            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Image image = Image.FromStream(originalStream))
                {
                    // Prepare encoder parameters for progressive JPEG with reduced quality.
                    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                                                             .First(c => c.FormatID == ImageFormat.Jpeg.Guid);

                    EncoderParameters encoderParams = new EncoderParameters(2);
                    // Quality = 70 (out of 100).
                    encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 70L);
                    // Enable progressive (interlaced) encoding.
                    encoderParams.Param[1] = new EncoderParameter(Encoder.ScanMethod,
                                                                  (long)EncoderValue.ScanMethodInterlaced);

                    // Save the compressed image.
                    string compressedPath = Path.Combine(artifactsDir, $"compressed_{index}.jpg");
                    image.Save(compressedPath, jpegCodec, encoderParams);

                    // Validate that the file was created.
                    if (!File.Exists(compressedPath))
                        throw new InvalidOperationException($"Failed to create compressed image: {compressedPath}");

                    index++;
                }
            }
        }

        // All done – the compressed JPEG images are stored in the Artifacts folder.
    }
}
