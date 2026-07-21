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
        // Prepare deterministic file names.
        const string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create sample images that will be embedded into the PDF.
        // -----------------------------------------------------------------
        string sampleImage1Path = Path.Combine(artifactsDir, "sample1.png");
        string sampleImage2Path = Path.Combine(artifactsDir, "sample2.png");

        CreateSampleImage(sampleImage1Path, 200, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2Path, 150, 150, Aspose.Drawing.Color.LightGreen);

        // -----------------------------------------------------------------
        // 2. Build a Word document, insert the images, and save it as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First image:");
        builder.InsertImage(sampleImage1Path);
        builder.Writeln();
        builder.Writeln("Second image:");
        builder.InsertImage(sampleImage2Path);

        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Load the PDF and extract each embedded image.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        var shapes = pdfDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                           .Where(s => s.HasImage)
                           .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No images were found in the PDF document.");

        int imageIndex = 0;
        foreach (var shape in shapes)
        {
            // Save the original image data to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into Aspose.Drawing.Bitmap.
                using (Bitmap bitmap = new Bitmap(originalStream))
                {
                    // Prepare JPEG encoder with 85% quality.
                    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                        .First(c => c.FormatID == ImageFormat.Jpeg.Guid);

                    EncoderParameters encoderParams = new EncoderParameters(1);
                    encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 85L);

                    // Save the bitmap as JPEG.
                    string outputJpegPath = Path.Combine(artifactsDir,
                        $"extracted_{imageIndex}.jpg");
                    bitmap.Save(outputJpegPath, jpegCodec, encoderParams);
                }
            }

            imageIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one JPEG file was created.
        // -----------------------------------------------------------------
        int jpegCount = Directory.GetFiles(artifactsDir, "extracted_*.jpg").Length;
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were created.");

        // The program finishes automatically.
    }

    // Helper method to create a deterministic PNG image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(backColor);
            bitmap.Save(filePath);
        }
    }
}
