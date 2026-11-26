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
        // Define paths
        string baseDir = Directory.GetCurrentDirectory();
        string docPath = Path.Combine(baseDir, "SampleDoc.docx");
        string bmpFolder = Path.Combine(baseDir, "ExtractedImages");
        string pngFolder = Path.Combine(baseDir, "ConvertedPngs");
        string checksumFile = Path.Combine(baseDir, "checksums.txt");

        // Ensure output directories exist
        Directory.CreateDirectory(bmpFolder);
        Directory.CreateDirectory(pngFolder);

        // Create sample BMP images
        string[] sampleBmpPaths = new string[2];
        for (int i = 0; i < sampleBmpPaths.Length; i++)
        {
            string bmpPath = Path.Combine(baseDir, $"sample{i + 1}.bmp");
            using (Bitmap bitmap = new Bitmap(200, 100))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.White);
                    g.FillRectangle(new SolidBrush(Color.FromArgb(255, 100 + i * 50, 150, 200)), 10, 10, 180, 80);
                }
                bitmap.Save(bmpPath, ImageFormat.Bmp);
            }
            sampleBmpPaths[i] = bmpPath;
        }

        // Create a Word document and insert BMP images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with BMP images:");
        foreach (string bmpPath in sampleBmpPaths)
        {
            builder.InsertImage(bmpPath);
            builder.Writeln(); // Add a line break after each image
        }
        doc.Save(docPath);

        // Load the document and extract embedded BMP images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Save the extracted image as BMP
                string extractedBmpPath = Path.Combine(bmpFolder, $"extracted-{imageIndex}.bmp");
                shape.ImageData.Save(extractedBmpPath);
                imageIndex++;
            }
        }

        // Verify that at least one BMP was extracted
        string[] extractedBmpFiles = Directory.GetFiles(bmpFolder, "*.bmp");
        if (extractedBmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP images were extracted from the document.");

        // Convert each BMP to lossless PNG
        foreach (string bmpFile in extractedBmpFiles)
        {
            using (Bitmap bmp = new Bitmap(bmpFile))
            {
                string pngFileName = Path.GetFileNameWithoutExtension(bmpFile) + ".png";
                string pngPath = Path.Combine(pngFolder, pngFileName);
                bmp.Save(pngPath, ImageFormat.Png);
            }
        }

        // Verify that at least one PNG was created
        string[] pngFiles = Directory.GetFiles(pngFolder, "*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were created during conversion.");

        // Generate checksum file for PNGs (SHA256)
        using (StreamWriter checksumWriter = new StreamWriter(checksumFile, false, Encoding.UTF8))
        {
            foreach (string pngPath in pngFiles)
            {
                byte[] fileBytes = File.ReadAllBytes(pngPath);
                byte[] hashBytes;
                using (SHA256 sha = SHA256.Create())
                {
                    hashBytes = sha.ComputeHash(fileBytes);
                }
                string hashString = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                string line = $"{Path.GetFileName(pngPath)}  {hashString}";
                checksumWriter.WriteLine(line);
            }
        }

        // Final validation: checksum file must exist and contain entries
        if (!File.Exists(checksumFile) || new FileInfo(checksumFile).Length == 0)
            throw new InvalidOperationException("Checksum file was not created or is empty.");
    }
}
