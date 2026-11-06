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
        // Define folders and files
        string inputFolder = "InputImages";
        string outputFolder = "OutputImages";
        string docPath = "DocumentWithImages.docx";
        string checksumFile = "checksums.txt";

        // Clean previous run data
        if (Directory.Exists(inputFolder)) Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder)) Directory.Delete(outputFolder, true);
        if (File.Exists(docPath)) File.Delete(docPath);
        if (File.Exists(checksumFile)) File.Delete(checksumFile);

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample BMP images
        for (int i = 1; i <= 2; i++)
        {
            string bmpPath = Path.Combine(inputFolder, $"sample{i}.bmp");
            using (Bitmap bitmap = new Bitmap(100, 100))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Distinct fill color for each image
                    Color fillColor = i == 1 ? Color.LightBlue : Color.LightGreen;
                    g.Clear(fillColor);
                }
                bitmap.Save(bmpPath, ImageFormat.Bmp);
            }
        }

        // Create a document and insert the BMP images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string bmpFile in Directory.GetFiles(inputFolder, "*.bmp"))
        {
            builder.InsertImage(bmpFile);
            builder.Writeln(); // separate images
        }
        doc.Save(docPath);

        // Load the document for extraction
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        var checksumBuilder = new StringBuilder();
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Save extracted image as PNG (lossless)
            string pngPath = Path.Combine(outputFolder, $"image{++imageIndex}.png");
            shape.ImageData.Save(pngPath); // format inferred from .png extension

            // Compute SHA256 checksum
            byte[] pngBytes = File.ReadAllBytes(pngPath);
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(pngBytes);
                string hashString = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                checksumBuilder.AppendLine($"{Path.GetFileName(pngPath)} {hashString}");
            }
        }

        // Validate that at least one PNG was created
        string[] pngFiles = Directory.GetFiles(outputFolder, "*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were generated.");

        // Write checksum file
        File.WriteAllText(checksumFile, checksumBuilder.ToString());

        // Validate checksum file
        if (!File.Exists(checksumFile) || new FileInfo(checksumFile).Length == 0)
            throw new InvalidOperationException("Checksum file was not created or is empty.");
    }
}
