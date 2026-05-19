using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for temporary files
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a deterministic JPEG image using Aspose.Drawing
        string jpegPath = Path.Combine(workDir, "sample.jpg");
        CreateSampleJpeg(jpegPath);

        // 2. Insert the JPEG into a Word document and save as PDF (simulating a scanned PDF)
        string pdfPath = Path.Combine(workDir, "sample.pdf");
        CreatePdfFromImage(jpegPath, pdfPath);

        // 3. Load the PDF document with Aspose.Words
        Document pdfDoc = new Document(pdfPath);

        // 4. Extract JPEG images from the PDF, correct EXIF orientation (simulated by rotating 90°), and save
        ExtractAndCorrectImages(pdfDoc, workDir);
    }

    // Creates a deterministic JPEG image file
    private static void CreateSampleJpeg(string filePath)
    {
        int width = 200;
        int height = 100;

        // Use Aspose.Drawing types explicitly to avoid ambiguity
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Use Aspose.Drawing.Font explicitly
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20))
            {
                g.DrawString("Sample", font, new SolidBrush(Color.Black), 10, 30);
            }

            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Inserts the image into a document and saves it as PDF
    private static void CreatePdfFromImage(string imagePath, string pdfPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);
    }

    // Extracts images, applies a rotation (as a placeholder for EXIF orientation correction), and saves them
    private static void ExtractAndCorrectImages(Document doc, string outputDir)
    {
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // Reset before reading

                // Load the image with Aspose.Drawing
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Rotate the image 90 degrees clockwise (simulating EXIF orientation correction)
                    using (Bitmap rotatedBitmap = RotateBitmap90Clockwise(originalBitmap))
                    {
                        string correctedPath = Path.Combine(outputDir, $"corrected_{imageIndex}.jpg");
                        rotatedBitmap.Save(correctedPath, ImageFormat.Jpeg);
                        imageIndex++;
                    }
                }
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the PDF document.");
    }

    // Rotates a bitmap 90 degrees clockwise using Aspose.Drawing
    private static Bitmap RotateBitmap90Clockwise(Bitmap source)
    {
        int newWidth = source.Height;
        int newHeight = source.Width;
        Bitmap rotated = new Bitmap(newWidth, newHeight);
        using (Graphics g = Graphics.FromImage(rotated))
        {
            // Move origin to center of new bitmap
            g.TranslateTransform(newWidth / 2f, newHeight / 2f);
            // Rotate 90 degrees
            g.RotateTransform(90);
            // Move origin back and draw the original bitmap
            g.TranslateTransform(-source.Width / 2f, -source.Height / 2f);
            g.DrawImage(source, 0, 0, source.Width, source.Height);
        }
        return rotated;
    }
}
