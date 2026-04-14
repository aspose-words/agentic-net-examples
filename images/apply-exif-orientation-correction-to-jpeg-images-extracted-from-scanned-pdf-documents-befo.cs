using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExifOrientationCorrectionExample
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic JPEG image using Aspose.Drawing.
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(jpegPath);

        // 2. Insert the JPEG into a PDF document.
        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        CreatePdfWithImage(jpegPath, pdfPath);

        // 3. Load the PDF and extract images.
        string outputDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(outputDir);
        ExtractImages(pdfPath, outputDir);

        // Validation: ensure at least one image was saved.
        if (Directory.GetFiles(outputDir, "*.*", SearchOption.TopDirectoryOnly).Length == 0)
            throw new InvalidOperationException("No images were extracted.");

        Console.WriteLine("Image extraction completed successfully.");
    }

    // Creates a deterministic JPEG image using Aspose.Drawing.
    private static void CreateSampleJpeg(string filePath)
    {
        int width = 200;
        int height = 100;

        // Create bitmap and draw simple text.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                    {
                        g.DrawString("Sample", font, brush, new Aspose.Drawing.PointF(10, 30));
                    }
                }

                // Save as JPEG.
                bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }

        // Validate that the file exists.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create JPEG at '{filePath}'.");
    }

    // Creates a PDF document containing the specified image.
    private static void CreatePdfWithImage(string imagePath, string pdfPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate PDF creation.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF at '{pdfPath}'.");
    }

    // Extracts images from a PDF and saves them to the output directory.
    private static void ExtractImages(string pdfPath, string outputDir)
    {
        Document pdfDoc = new Document(pdfPath);
        NodeCollection shapes = pdfDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine file extension based on image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outputPath = Path.Combine(outputDir, $"extracted_{imageIndex}{extension}");

            // Save the image data directly.
            shape.ImageData.Save(outputPath);
            imageIndex++;
        }
    }
}
