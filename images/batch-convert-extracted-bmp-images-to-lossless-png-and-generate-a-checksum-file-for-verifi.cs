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
        // Prepare folders
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Step 1: Create sample BMP images
        CreateSampleBmp(Path.Combine(inputFolder, "sample1.bmp"), 100, 100, Aspose.Drawing.Color.LightBlue);
        CreateSampleBmp(Path.Combine(inputFolder, "sample2.bmp"), 120, 80, Aspose.Drawing.Color.LightCoral);

        // Step 2: Insert BMP images into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string bmpPath in Directory.GetFiles(inputFolder, "*.bmp"))
        {
            builder.InsertImage(bmpPath);
            builder.Writeln(); // separate images
        }
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        doc.Save(docPath);

        // Step 3: Extract images from the document back to BMP files (overwrite existing)
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                string extractedPath = Path.Combine(inputFolder, $"extracted_{extractedIndex}.bmp");
                shape.ImageData.Save(extractedPath);
                extractedIndex++;
            }
        }

        // Step 4: Batch convert BMP images to lossless PNG
        string[] bmpFiles = Directory.GetFiles(inputFolder, "*.bmp");
        if (bmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP files found for conversion.");

        foreach (string bmpFile in bmpFiles)
        {
            using (Bitmap bitmap = new Bitmap(bmpFile))
            {
                string pngFileName = Path.GetFileNameWithoutExtension(bmpFile) + ".png";
                string pngPath = Path.Combine(outputFolder, pngFileName);
                bitmap.Save(pngPath, ImageFormat.Png);
            }
        }

        // Validate PNG output
        string[] pngFiles = Directory.GetFiles(outputFolder, "*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("PNG conversion produced no files.");

        // Step 5: Generate checksum file for verification
        string checksumFilePath = Path.Combine(outputFolder, "checksums.txt");
        using (StreamWriter writer = new StreamWriter(checksumFilePath, false, Encoding.UTF8))
        {
            foreach (string pngFile in pngFiles)
            {
                string checksum = ComputeSha256(pngFile);
                string line = $"{Path.GetFileName(pngFile)} {checksum}";
                writer.WriteLine(line);
            }
        }

        // Validate checksum file
        if (!File.Exists(checksumFilePath))
            throw new InvalidOperationException("Checksum file was not created.");
    }

    private static void CreateSampleBmp(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(fillColor);
            }
            bitmap.Save(filePath, ImageFormat.Bmp);
        }
    }

    private static string ComputeSha256(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hash = sha256.ComputeHash(stream);
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hash)
                sb.Append(b.ToString("x2"));
            return sb.ToString();
        }
    }
}
