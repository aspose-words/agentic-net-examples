using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Base directory for all artifacts
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(baseDir);

        // Input folder for sample BMP images
        string inputDir = Path.Combine(baseDir, "InputImages");
        Directory.CreateDirectory(inputDir);

        // Output folder for converted PNG images
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(outputDir);

        // Path for the checksum file
        string checksumFile = Path.Combine(baseDir, "checksums.txt");

        // 1. Create deterministic BMP sample images
        CreateSampleBmpImages(inputDir, 3);

        // 2. Create a Word document and insert the BMP images
        string docPath = Path.Combine(baseDir, "Sample.docx");
        CreateDocumentWithImages(docPath, inputDir);

        // 3. Load the document, extract images, convert to PNG, and generate checksums
        ConvertImagesToPng(docPath, outputDir, checksumFile);
    }

    // Creates a given number of BMP files with simple graphics
    private static void CreateSampleBmpImages(string folder, int count)
    {
        for (int i = 0; i < count; i++)
        {
            string filePath = Path.Combine(folder, $"sample{i}.bmp");
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100 + i * 20, 100 + i * 20))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(
                        Aspose.Drawing.Color.FromArgb(255, 50 + i * 50, 100, 150)))
                    {
                        g.FillRectangle(brush, 10, 10, bitmap.Width - 20, bitmap.Height - 20);
                    }
                }
                bitmap.Save(filePath);
            }

            if (!File.Exists(filePath))
                throw new InvalidOperationException($"Failed to create BMP image: {filePath}");
        }
    }

    // Builds a Word document and inserts all BMP files from the specified folder
    private static void CreateDocumentWithImages(string docPath, string imagesFolder)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string bmpFile in Directory.GetFiles(imagesFolder, "*.bmp"))
        {
            builder.InsertParagraph();
            builder.InsertImage(bmpFile);
        }

        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save document: {docPath}");
    }

    // Extracts all images from the document, converts each to PNG, and writes SHA256 checksums
    private static void ConvertImagesToPng(string docPath, string outputFolder, string checksumFilePath)
    {
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        StringBuilder checksumBuilder = new StringBuilder();

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Get the raw image bytes from the shape
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into an Aspose.Drawing.Bitmap
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(ms))
            {
                string pngPath = Path.Combine(outputFolder, $"image{imageIndex}.png");
                bitmap.Save(pngPath, ImageFormat.Png);

                if (!File.Exists(pngPath))
                    throw new InvalidOperationException($"Failed to save PNG image: {pngPath}");

                // Compute SHA256 checksum of the newly saved PNG
                string checksum = ComputeSha256(pngPath);
                checksumBuilder.AppendLine($"{Path.GetFileName(pngPath)} {checksum}");

                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found for conversion.");

        // Write checksum file
        File.WriteAllText(checksumFilePath, checksumBuilder.ToString());

        if (!File.Exists(checksumFilePath))
            throw new InvalidOperationException("Checksum file was not created.");
    }

    // Computes SHA256 hash of a file and returns it as a lowercase hex string
    private static string ComputeSha256(string filePath)
    {
        using (SHA256 sha256 = SHA256.Create())
        using (FileStream stream = File.OpenRead(filePath))
        {
            byte[] hash = sha256.ComputeHash(stream);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }
    }
}
