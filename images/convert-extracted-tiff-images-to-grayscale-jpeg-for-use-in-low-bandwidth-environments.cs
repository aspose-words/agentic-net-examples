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
        // Prepare a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample TIFF image.
        string tiffPath = Path.Combine(artifactsDir, "sample.tiff");
        CreateSampleTiff(tiffPath);

        // 2. Insert the TIFF image into a Word document.
        string docPath = Path.Combine(artifactsDir, "docWithTiff.docx");
        InsertImageIntoDocument(tiffPath, docPath);

        // 3. Extract the image(s) from the document and convert each to a grayscale JPEG.
        ConvertExtractedImagesToGrayscaleJpeg(docPath, artifactsDir);
    }

    // Creates a deterministic 200x200 TIFF image with a red rectangle.
    private static void CreateSampleTiff(string filePath)
    {
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
                {
                    g.FillRectangle(brush, 20, 20, 160, 160);
                }
            }

            bitmap.Save(filePath, ImageFormat.Tiff);
        }

        if (!File.Exists(filePath))
            throw new Exception("Failed to create sample TIFF image.");
    }

    // Inserts the provided image file into a new Word document.
    private static void InsertImageIntoDocument(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception("Failed to save document with TIFF image.");
    }

    // Extracts all images from the document, converts each to grayscale,
    // and saves the result as a JPEG file.
    private static void ConvertExtractedImagesToGrayscaleJpeg(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image into a bitmap for pixel manipulation.
                using (Bitmap bitmap = new Bitmap(imageStream))
                {
                    // Convert the bitmap to grayscale.
                    ConvertBitmapToGrayscale(bitmap);

                    // Prepare the output JPEG path.
                    string jpegPath = Path.Combine(outputDir, $"grayscale_{imageIndex}.jpg");

                    // Save the grayscale bitmap as JPEG.
                    bitmap.Save(jpegPath, ImageFormat.Jpeg);

                    if (!File.Exists(jpegPath))
                        throw new Exception($"Failed to save JPEG image: {jpegPath}");

                    imageIndex++;
                }
            }
        }

        if (imageIndex == 0)
            throw new Exception("No images were found and converted.");
    }

    // Performs a per‑pixel grayscale conversion using the standard luminance formula.
    private static void ConvertBitmapToGrayscale(Bitmap bitmap)
    {
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);
                int gray = (int)(0.3 * original.R + 0.59 * original.G + 0.11 * original.B);
                Aspose.Drawing.Color grayColor = Aspose.Drawing.Color.FromArgb(gray, gray, gray);
                bitmap.SetPixel(x, y, grayColor);
            }
        }
    }
}
