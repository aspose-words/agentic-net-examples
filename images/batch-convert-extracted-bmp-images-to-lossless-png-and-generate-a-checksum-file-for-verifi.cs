using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
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
        // Directories for sample input BMPs and output PNGs
        string inputDir = "InputImages";
        string outputDir = "OutputImages";
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Step 1: Create deterministic sample BMP images using Aspose.Drawing
        for (int i = 0; i < 3; i++)
        {
            string bmpPath = Path.Combine(inputDir, $"sample{i}.bmp");
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    // Fill each bitmap with a distinct color
                    g.Clear(Aspose.Drawing.Color.FromArgb(50 + i * 50, 100 + i * 30, 150 + i * 20));
                }

                // Save as BMP
                bitmap.Save(bmpPath, Aspose.Drawing.Imaging.ImageFormat.Bmp);
            }
        }

        // Step 2: Create a Word document and insert the BMP images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 3; i++)
        {
            string bmpPath = Path.Combine(inputDir, $"sample{i}.bmp");
            // Insert the BMP file; the image type is preserved when possible
            builder.InsertImage(bmpPath);
            builder.Writeln(); // separate images with a line break
        }

        string docPath = "SampleDocument.docx";
        doc.Save(docPath, SaveFormat.Docx);

        // Step 3: Load the document and batch convert extracted images to PNG
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int convertedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0; // reset before reading

                // Load the image via Aspose.Drawing.Bitmap (supports BMP, PNG, etc.)
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgStream))
                {
                    string pngFileName = $"image{convertedCount}.png";
                    string pngPath = Path.Combine(outputDir, pngFileName);

                    // Save as lossless PNG
                    bitmap.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);
                    convertedCount++;
                }
            }
        }

        // Validation: ensure at least one PNG was produced
        if (convertedCount == 0)
            throw new InvalidOperationException("No images were found for conversion.");

        // Step 4: Generate checksum file for the PNG images
        string checksumFile = Path.Combine(outputDir, "checksums.txt");
        using (StreamWriter writer = new StreamWriter(checksumFile, false, Encoding.UTF8))
        {
            for (int i = 0; i < convertedCount; i++)
            {
                string pngPath = Path.Combine(outputDir, $"image{i}.png");
                byte[] pngBytes = File.ReadAllBytes(pngPath);
                using (SHA256 sha256 = SHA256.Create())
                {
                    byte[] hash = sha256.ComputeHash(pngBytes);
                    string hashHex = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                    writer.WriteLine($"{Path.GetFileName(pngPath)} {hashHex}");
                }
            }
        }

        // Final validation: checksum file must exist and contain entries
        if (!File.Exists(checksumFile) || new FileInfo(checksumFile).Length == 0)
            throw new InvalidOperationException("Checksum file was not created correctly.");
    }
}
