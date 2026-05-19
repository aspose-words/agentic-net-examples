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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic JPEG image.
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(jpegPath);

        // 2. Create a DOCX document and insert the JPEG image.
        string docPath = Path.Combine(artifactsDir, "Document.docx");
        CreateDocumentWithImages(docPath, jpegPath);

        // 3. Load the document, apply grayscale to JPEG images, and save them.
        ExtractAndGrayscaleImages(docPath, artifactsDir);
    }

    // Creates a simple 100x100 JPEG image with a red rectangle.
    private static void CreateSampleJpeg(string filePath)
    {
        using (var bitmap = new Bitmap(100, 100))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (var pen = new Pen(Color.Red, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Builds a document that contains two copies of the JPEG image.
    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.InsertImage(imagePath);

        doc.Save(docPath);
    }

    // Loads the document, sets GrayScale on JPEG images, and extracts them to files.
    private static void ExtractAndGrayscaleImages(string docPath, string outputDir)
    {
        var doc = new Document(docPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Apply grayscale only to JPEG images.
            if (shape.ImageData.ImageType == ImageType.Jpeg)
                shape.ImageData.GrayScale = true;

            // Determine proper file extension and save the image.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outPath = Path.Combine(outputDir, $"Extracted_{imageIndex}{extension}");
            shape.ImageData.Save(outPath);
            imageIndex++;
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
