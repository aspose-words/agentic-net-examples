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
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a large PNG image (> 2 MB) filled with random colors.
        // -----------------------------------------------------------------
        string largeImagePath = Path.Combine(artifactsDir, "large.png");
        CreateLargePng(largeImagePath, 3000, 3000); // 3000×3000 pixels with random data.

        // -----------------------------------------------------------------
        // 2. Build a sample DOCX containing the large image twice.
        // -----------------------------------------------------------------
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        CreateDocumentWithImage(inputDocPath, largeImagePath);

        // -----------------------------------------------------------------
        // 3. Load the document and replace images larger than 2 MB with JPEGs.
        // -----------------------------------------------------------------
        Document doc = new Document(inputDocPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        const long sizeThreshold = 2L * 1024 * 1024; // 2 MB
        int replacedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            byte[] originalBytes = shape.ImageData.ImageBytes;
            if (originalBytes == null || originalBytes.Length <= sizeThreshold)
                continue;

            // Convert the original image to JPEG using a memory stream.
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            using (Bitmap bitmap = new Bitmap(originalStream))
            using (MemoryStream jpegStream = new MemoryStream())
            {
                bitmap.Save(jpegStream, ImageFormat.Jpeg);
                jpegStream.Position = 0; // Reset before feeding to SetImage.
                shape.ImageData.SetImage(jpegStream);
                replacedCount++;
            }
        }

        if (replacedCount == 0)
            throw new InvalidOperationException("No images larger than 2 MB were found to replace.");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");
        doc.Save(outputDocPath, SaveFormat.Docx);

        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }

    // Creates a PNG image of the specified size filled with random colors.
    private static void CreateLargePng(string filePath, int width, int height)
    {
        Random rnd = new Random();
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill each pixel with a random color to avoid strong compression.
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    int r = rnd.Next(256);
                    int g = rnd.Next(256);
                    int b = rnd.Next(256);
                    bitmap.SetPixel(x, y, Color.FromArgb(r, g, b));
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a DOCX file that contains the specified image twice.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.InsertImage(imagePath);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);

        doc.Save(docPath, SaveFormat.Docx);
    }
}
