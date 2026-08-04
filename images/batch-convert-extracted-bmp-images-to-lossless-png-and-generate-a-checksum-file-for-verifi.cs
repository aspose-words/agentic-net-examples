using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Set up directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create deterministic BMP sample images.
        for (int i = 1; i <= 3; i++)
        {
            string bmpPath = Path.Combine(inputDir, $"sample{i}.bmp");
            using (Bitmap bmp = new Bitmap(100, 100))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.FromArgb(50 * i, 80 * i, 120));
                bmp.Save(bmpPath, ImageFormat.Bmp);
            }
        }

        // 2. Insert the BMP images into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string bmpFile in Directory.GetFiles(inputDir, "*.bmp"))
        {
            builder.InsertParagraph();
            builder.InsertImage(bmpFile);
        }

        string docPath = Path.Combine(baseDir, "ImagesDoc.docx");
        doc.Save(docPath);

        // 3. Load the document and extract each image, converting it to PNG.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int pngIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Get the original image bytes (whatever format they are stored in).
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Load the bytes into a bitmap and re‑save as PNG.
            using (MemoryStream srcStream = new MemoryStream(originalBytes))
            using (Bitmap bmp = new Bitmap(srcStream))
            using (MemoryStream pngStream = new MemoryStream())
            {
                bmp.Save(pngStream, ImageFormat.Png);
                pngStream.Position = 0;

                string pngPath = Path.Combine(outputDir, $"image{pngIndex}.png");
                using (FileStream outFile = new FileStream(pngPath, FileMode.Create, FileAccess.Write))
                {
                    pngStream.CopyTo(outFile);
                }

                pngIndex++;
            }
        }

        // Validate that PNG files were created.
        string[] pngFiles = Directory.GetFiles(outputDir, "*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were generated.");

        // 4. Generate a SHA‑256 checksum file for the PNGs.
        StringBuilder checksumBuilder = new StringBuilder();
        using (SHA256 sha256 = SHA256.Create())
        {
            foreach (string pngFile in pngFiles.OrderBy(f => f))
            {
                byte[] fileBytes = File.ReadAllBytes(pngFile);
                byte[] hashBytes = sha256.ComputeHash(fileBytes);
                string hashString = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                string fileName = Path.GetFileName(pngFile);
                checksumBuilder.AppendLine($"{fileName} {hashString}");
            }
        }

        string checksumPath = Path.Combine(baseDir, "checksums.txt");
        File.WriteAllText(checksumPath, checksumBuilder.ToString());

        // Final validation of checksum file.
        if (!File.Exists(checksumPath) || new FileInfo(checksumPath).Length == 0)
            throw new InvalidOperationException("Checksum file was not created correctly.");
    }
}
