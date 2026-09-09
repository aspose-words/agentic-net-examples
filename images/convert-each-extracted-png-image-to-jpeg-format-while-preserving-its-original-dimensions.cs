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
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic PNG sample image.
        string pngPath = Path.Combine(artifactsDir, "sample.png");
        CreateSamplePng(pngPath, 200, 100);

        // 2. Insert the PNG into a Word document.
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        InsertImageIntoDocument(pngPath, docPath);

        // 3. Load the document and convert each extracted PNG to JPEG.
        ConvertExtractedPngsToJpeg(docPath, artifactsDir);

        // 4. Validation – ensure at least one JPEG was produced.
        int jpegCount = Directory.GetFiles(artifactsDir, "*.jpg").Length;
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG files were created.");

        Console.WriteLine($"Conversion completed. {jpegCount} JPEG file(s) saved to '{artifactsDir}'.");
    }

    // Creates a simple PNG image using Aspose.Drawing.
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                g.FillRectangle(brush, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Inserts the given image file into a new document and saves it.
    private static void InsertImageIntoDocument(string imagePath, string docPath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts PNG images from the document, converts each to JPEG, and saves them.
    private static void ConvertExtractedPngsToJpeg(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .Cast<Shape>()
                        .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Png)
                        .ToList();

        int index = 0;
        foreach (var shape in shapes)
        {
            // Save the shape's image data to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset before reading.

                // Load the image via Aspose.Drawing.
                using (Image img = Image.FromStream(ms))
                {
                    string jpegPath = Path.Combine(outputDir, $"extracted_{index}.jpg");
                    img.Save(jpegPath, ImageFormat.Jpeg);
                }
            }
            index++;
        }
    }
}
