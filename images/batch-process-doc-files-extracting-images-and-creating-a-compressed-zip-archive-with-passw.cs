using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    // Password used for encrypting the final ZIP archive
    private const string ZipPassword = "StrongPassword123";

    // Entry point
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Create sample DOCX files that contain the image
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc2.docx"), sampleImagePath);

        // 3. Batch process DOCX files: extract images
        int totalExtracted = ExtractImagesFromDocuments(inputDocsDir, imagesDir);
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // 4. Create a ZIP archive of the extracted images
        string zipPath = Path.Combine(outputDir, "ImagesArchive.zip");
        ZipFile.CreateFromDirectory(imagesDir, zipPath, CompressionLevel.Optimal, false);
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the ZIP archive.");

        // 5. Encrypt the ZIP archive with a password (AES)
        string encryptedZipPath = Path.Combine(outputDir, "ImagesArchive.zip.enc");
        EncryptFileWithPassword(zipPath, encryptedZipPath, ZipPassword);
        if (!File.Exists(encryptedZipPath))
            throw new InvalidOperationException("Failed to encrypt the ZIP archive.");

        Console.WriteLine("Batch processing completed successfully.");
        Console.WriteLine($"Extracted images count: {totalExtracted}");
        Console.WriteLine($"Encrypted ZIP archive created at: {encryptedZipPath}");
    }

    // Creates a simple white PNG image with a black rectangle for demonstration
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple black rectangle
                using (Pen pen = new Pen(Color.Black, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at {filePath}.");
    }

    // Generates a DOCX file containing the specified image
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create sample document at {docPath}.");
    }

    // Extracts all images from DOCX files in the source folder into the destination folder
    private static int ExtractImagesFromDocuments(string sourceFolder, string destinationFolder)
    {
        int extractedCount = 0;
        string[] docFiles = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.TopDirectoryOnly);

        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageExt = shape.ImageData.ImageType.ToString().ToLower(); // e.g., png, jpeg
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}.{imageExt}";
                    string imagePath = Path.Combine(destinationFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    if (!File.Exists(imagePath))
                        throw new InvalidOperationException($"Failed to save extracted image to {imagePath}.");

                    extractedCount++;
                    imageIndex++;
                }
            }
        }

        return extractedCount;
    }

    // Encrypts a file using AES (CBC) with a password-derived key
    private static void EncryptFileWithPassword(string inputFile, string outputFile, string password)
    {
        // Derive a 256-bit key from the password using SHA256
        byte[] key;
        using (SHA256 sha = SHA256.Create())
        {
            key = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
        }

        // Generate a random IV and prepend it to the encrypted file
        using (Aes aes = Aes.Create())
        {
            aes.Key = key;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;
            aes.GenerateIV();

            using (FileStream fsInput = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
            using (FileStream fsOutput = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
            {
                // Write IV first
                fsOutput.Write(aes.IV, 0, aes.IV.Length);

                using (CryptoStream cryptoStream = new CryptoStream(fsOutput, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    fsInput.CopyTo(cryptoStream);
                }
            }
        }
    }
}
