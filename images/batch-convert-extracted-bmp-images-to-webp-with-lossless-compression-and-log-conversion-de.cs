using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;          // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging; // For ImageFormat

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create sample BMP images and insert them into a Word document.
        // -----------------------------------------------------------------
        const int imageCount = 3;
        string[] bmpPaths = new string[imageCount];

        for (int i = 0; i < imageCount; i++)
        {
            // Create a deterministic 100x100 bitmap with a solid color.
            using (Bitmap bitmap = new Bitmap(100, 100))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.FromArgb(50 * i, 100 + 30 * i, 150 + 20 * i));
                string bmpPath = Path.Combine(artifactsDir, $"sample{i + 1}.bmp");
                bitmap.Save(bmpPath, ImageFormat.Bmp);
                bmpPaths[i] = bmpPath;
            }
        }

        // Build a document and insert the BMP images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string bmpPath in bmpPaths)
        {
            builder.InsertParagraph(); // separate images
            builder.InsertImage(bmpPath);
        }

        // Save the source document.
        string sourceDocPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        doc.Save(sourceDocPath);

        // ---------------------------------------------------------------
        // 2. Load the document and convert each image to lossless WebP.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int conversionIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue; // Skip shapes without image data.

            // Obtain original image size (in bytes).
            long originalSize;
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalSize = originalStream.Length;
            }

            // Define output WebP file name.
            string webpPath = Path.Combine(artifactsDir, $"ConvertedImage{conversionIndex}.webp");

            // Save the image as WebP. Lossless compression is the default when no quality is specified.
            shape.ImageData.Save(webpPath);

            // Obtain new image size (in bytes).
            long newSize = new FileInfo(webpPath).Length;

            // Log conversion details.
            Console.WriteLine(
                $"Conversion {conversionIndex + 1}: Original ({originalSize} bytes) → WebP ({newSize} bytes) – saved as '{Path.GetFileName(webpPath)}'.");

            conversionIndex++;
        }

        // Validate that at least one image was converted.
        if (conversionIndex == 0)
            throw new InvalidOperationException("No images were found for conversion.");

        Console.WriteLine("Batch conversion completed successfully.");
    }
}
