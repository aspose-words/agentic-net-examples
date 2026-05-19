using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Deterministic file and folder names.
        const string sampleImagePath = "sample.png";
        const string pdfPath = "sample.pdf";
        const string outputFolder = "ExtractedImages";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 20, 20, imgWidth - 40, imgHeight - 40);
                }
            }
            // Save the bitmap as PNG – this file will be inserted into the document.
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a Word document, insert the image, and save as PDF.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);

        // Save as PDF with JPEG compression (quality 85) – this affects JPEG images inside the PDF.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 85
        };
        doc.Save(pdfPath, pdfSaveOptions);

        // -------------------------------------------------
        // 3. Load the PDF and extract embedded images.
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        var shapeImages = pdfDoc.GetChildNodes(NodeType.Shape, true)
                                .Cast<Shape>()
                                .Where(s => s.HasImage)
                                .ToList();

        if (!shapeImages.Any())
            throw new InvalidOperationException("No images were found in the PDF document.");

        int imageIndex = 0;
        foreach (var shape in shapeImages)
        {
            // Save the original image data to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Image originalImage = Image.FromStream(imgStream))
                {
                    // Convert to Bitmap to enable JPEG saving with quality.
                    using (Bitmap bmp = new Bitmap(originalImage))
                    {
                        string outFile = Path.Combine(outputFolder, $"image_{imageIndex}.jpg");

                        // Prepare JPEG codec and encoder parameters (quality = 85).
                        ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                            .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                        EncoderParameters encoderParams = new EncoderParameters(1);
                        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 85L);

                        // Save as JPEG with the specified quality.
                        bmp.Save(outFile, jpegCodec, encoderParams);
                    }
                }
            }
            imageIndex++;
        }

        // -------------------------------------------------
        // 4. Validation – ensure at least one JPEG file was created.
        // -------------------------------------------------
        int jpegCount = Directory.GetFiles(outputFolder, "*.jpg").Length;
        if (jpegCount == 0)
            throw new InvalidOperationException("Failed to create any JPEG images.");
    }
}
