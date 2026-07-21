using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image.
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleJpegPath);

        // 2. Create a DOCX document and insert the JPEG image.
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, sampleJpegPath);

        // 3. Load the document and process JPEG images.
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int jpegCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Convert to grayscale using Aspose.Drawing.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap original = new Bitmap(ms))
            using (Bitmap grayBitmap = new Bitmap(original.Width, original.Height))
            {
                for (int y = 0; y < original.Height; y++)
                {
                    for (int x = 0; x < original.Width; x++)
                    {
                        Color pixel = original.GetPixel(x, y);
                        int gray = (pixel.R + pixel.G + pixel.B) / 3;
                        Color grayColor = Color.FromArgb(pixel.A, gray, gray, gray);
                        grayBitmap.SetPixel(x, y, grayColor);
                    }
                }

                string outputImagePath = Path.Combine(artifactsDir, $"Extracted_Gray_{jpegCount}.jpg");
                grayBitmap.Save(outputImagePath, ImageFormat.Jpeg);
                jpegCount++;
            }
        }

        // Validation: ensure at least one JPEG image was processed.
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were found in the document.");

        // Optional: clean up temporary files (commented out to keep artifacts for inspection).
        // File.Delete(sampleJpegPath);
        // File.Delete(docPath);
    }

    private static void CreateSampleJpeg(string filePath)
    {
        // Create a 200x200 bitmap, fill with a solid color, and save as JPEG.
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.CornflowerBlue);
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
