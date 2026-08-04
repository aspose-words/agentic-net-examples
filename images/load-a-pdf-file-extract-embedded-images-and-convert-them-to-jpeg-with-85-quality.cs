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
        // Prepare file paths.
        string workDir = Directory.GetCurrentDirectory();
        string sampleImagePath = Path.Combine(workDir, "sample.png");
        string pdfPath = Path.Combine(workDir, "sample.pdf");
        string outputDir = Path.Combine(workDir, "ExtractedImages");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (PNG).
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle.
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document, insert the image, and save as PDF.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 3. Load the PDF and extract embedded images.
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the shape's image data to a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image with Aspose.Drawing.
                using (Image img = Image.FromStream(imgStream))
                {
                    // Prepare JPEG encoder with 85% quality.
                    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                        .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                    EncoderParameters encoderParams = new EncoderParameters(1);
                    encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 85L);

                    // Save as JPEG.
                    string outFile = Path.Combine(outputDir, $"image_{imageIndex}.jpg");
                    img.Save(outFile, jpegCodec, encoderParams);
                }
            }

            imageIndex++;
        }

        // -------------------------------------------------
        // 4. Validation – ensure at least one image was extracted.
        // -------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the PDF.");

        // The program finishes automatically.
    }
}
