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
        // Ensure output directory exists
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample PNG image
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSamplePng(sampleImagePath);

        // 2. Create a PDF document, insert the image, and save
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        CreatePdfWithImage(sampleImagePath, pdfPath);

        // 3. Load the PDF and extract embedded images as JPEG with 85% quality
        ExtractImagesFromPdf(pdfPath, outputDir);
    }

    private static void CreateSamplePng(string filePath)
    {
        int width = 200;
        int height = 100;

        // Create bitmap using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain graphics object
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background
                g.Clear(Color.White);

                // Draw rectangle
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }

                // Draw filled ellipse
                using (Brush brush = new SolidBrush(Color.Red))
                {
                    g.FillEllipse(brush, 50, 25, 100, 50);
                }
            }

            // Save as PNG
            bitmap.Save(filePath, ImageFormat.Png);
        }

        // Validate that the image file was created
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
    }

    private static void CreatePdfWithImage(string imagePath, string pdfPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, Aspose.Words.SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF at '{pdfPath}'.");
    }

    private static void ExtractImagesFromPdf(string pdfPath, string outputDir)
    {
        Document pdfDoc = new Document(pdfPath);
        NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 1;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a temporary file
            string tempPath = Path.Combine(outputDir, $"temp-{imageIndex}.png");
            shape.ImageData.Save(tempPath);

            // Load the temporary image and re‑save as JPEG with 85% quality
            string jpegPath = Path.Combine(outputDir, $"extracted-{imageIndex}.jpg");
            using (Bitmap bitmap = new Bitmap(tempPath))
            {
                // Prepare encoder parameters for JPEG quality
                ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                    .First(codec => codec.FormatID == ImageFormat.Jpeg.Guid);

                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 85L);

                bitmap.Save(jpegPath, jpegCodec, encoderParams);
            }

            // Clean up temporary file
            if (File.Exists(tempPath))
                File.Delete(tempPath);

            imageIndex++;
        }

        if (imageIndex == 1)
            throw new InvalidOperationException("No images were extracted from the PDF.");

        // Validate that at least one JPEG file exists
        string[] extractedFiles = Directory.GetFiles(outputDir, "extracted-*.jpg");
        if (extractedFiles.Length == 0)
            throw new InvalidOperationException("Extraction reported success but no JPEG files were found.");
    }
}
