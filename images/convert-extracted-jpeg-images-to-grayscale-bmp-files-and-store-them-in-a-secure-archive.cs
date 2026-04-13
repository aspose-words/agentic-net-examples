using System;
using System.IO;
using System.IO.Compression;
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

        // 1. Create a deterministic JPEG image using Aspose.Drawing
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(jpegPath, 200, 200);

        // 2. Create a Word document and insert the JPEG image
        string docPath = Path.Combine(artifactsDir, "Document.docx");
        CreateDocumentWithImage(docPath, jpegPath);

        // 3. Load the document and extract JPEG images
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        string tempFolder = Path.Combine(artifactsDir, "TempImages");
        Directory.CreateDirectory(tempFolder);

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;

            // Process only JPEG images
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Get raw image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load bytes into Aspose.Drawing.Bitmap
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // reset before use
                using (Bitmap bitmap = new Bitmap(ms))
                {
                    // Convert to grayscale
                    ConvertToGrayscale(bitmap);

                    // Save as BMP
                    string bmpFileName = Path.Combine(tempFolder, $"image_{imageIndex}.bmp");
                    bitmap.Save(bmpFileName, ImageFormat.Bmp);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one BMP was created
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were extracted and converted.");

        // 4. Store the BMP files in a secure archive (ZIP)
        string zipPath = Path.Combine(artifactsDir, "ImagesArchive.zip");
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        // Resolve ambiguity by using the fully qualified System.IO.Compression.CompressionLevel
        ZipFile.CreateFromDirectory(tempFolder, zipPath, System.IO.Compression.CompressionLevel.Optimal, false);

        // Validate that the ZIP archive was created
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the ZIP archive.");

        // Cleanup temporary folder
        Directory.Delete(tempFolder, true);
    }

    // Creates a deterministic JPEG image file using Aspose.Drawing
    private static void CreateSampleJpeg(string filePath, int width, int height)
    {
        // Explicitly use Aspose.Drawing types to avoid ambiguity
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.White);
            // Draw a simple red rectangle
            using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                g.FillRectangle(brush, width / 4, height / 4, width / 2, height / 2);
            }
            // Save as JPEG
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }

    // Creates a Word document and inserts the given image
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Converts the provided bitmap to grayscale in-place
    private static void ConvertToGrayscale(Bitmap bitmap)
    {
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Aspose.Drawing.Color original = bitmap.GetPixel(x, y);
                // Compute luminance using Rec. 601 luma formula
                int lum = (int)(0.299 * original.R + 0.587 * original.G + 0.114 * original.B);
                Aspose.Drawing.Color gray = Aspose.Drawing.Color.FromArgb(lum, lum, lum);
                bitmap.SetPixel(x, y, gray);
            }
        }
    }
}
