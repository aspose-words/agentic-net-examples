using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Entry point
    public static void Main()
    {
        // Define folders
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageProcessing");
        string docsDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "ImagesArchive.zip");
        string zipPassword = "Secret123";

        // Ensure clean environment
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(imagesDir);

        // Create sample DOCX files with images
        for (int i = 1; i <= 3; i++)
        {
            string imagePath = Path.Combine(baseDir, $"SampleImage{i}.png");
            CreateSampleImage(imagePath, 200 + i * 20, 150 + i * 10, $"Img{i}");
            CreateSampleDocWithImage(docsDir, $"Document{i}.docx", imagePath);
        }

        // Process each DOCX file: extract images
        var docFiles = Directory.GetFiles(docsDir, "*.docx");
        int totalExtracted = 0;
        foreach (var docFile in docFiles)
        {
            totalExtracted += ExtractImagesFromDoc(docFile, imagesDir);
        }

        // Validate that images were extracted
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Create ZIP archive and encrypt it with a password
        CreatePasswordProtectedZip(imagesDir, zipPath, zipPassword);

        // Verify that the ZIP file exists
        if (!File.Exists(zipPath))
            throw new FileNotFoundException("Failed to create the ZIP archive.");

        // Example completed
        Console.WriteLine($"Processed {docFiles.Length} documents, extracted {totalExtracted} images.");
        Console.WriteLine($"Encrypted ZIP archive created at: {zipPath}");
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height, string text)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20, Aspose.Drawing.FontStyle.Bold))
            {
                graphics.DrawString(text, font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black), new Aspose.Drawing.PointF(10, 10));
            }
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
    }

    // Creates a DOCX file containing the specified image
    private static void CreateSampleDocWithImage(string docsFolder, string docName, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document containing image: {Path.GetFileName(imagePath)}");
        builder.InsertImage(imagePath);
        string fullPath = Path.Combine(docsFolder, docName);
        doc.Save(fullPath);
    }

    // Extracts all images from a document and saves them to the target folder
    private static int ExtractImagesFromDoc(string docPath, string targetFolder)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"Img_{Path.GetFileNameWithoutExtension(docPath)}_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(targetFolder, imageFileName);
                shape.ImageData.Save(imageFullPath);
                imageIndex++;
            }
        }
        return imageIndex;
    }

    // Creates a ZIP archive of the source folder and encrypts it with AES using the supplied password
    private static void CreatePasswordProtectedZip(string sourceFolder, string zipFilePath, string password)
    {
        // Generate a temporary ZIP file name that does not already exist
        string tempZip = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.zip");

        // Step 1: Create a temporary ZIP file without encryption
        ZipFile.CreateFromDirectory(sourceFolder, tempZip, CompressionLevel.Optimal, false);

        // Step 2: Encrypt the ZIP file using AES (CBC) with a key derived from the password
        byte[] salt = GenerateRandomBytes(16);
        using (Aes aes = Aes.Create())
        {
            var key = new Rfc2898DeriveBytes(password, salt, 100_000, HashAlgorithmName.SHA256);
            aes.Key = key.GetBytes(aes.KeySize / 8);
            aes.IV = GenerateRandomBytes(aes.BlockSize / 8);
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;

            using (FileStream fsOut = new FileStream(zipFilePath, FileMode.Create, FileAccess.Write))
            {
                // Write salt and IV at the beginning for later decryption
                fsOut.Write(salt, 0, salt.Length);
                fsOut.Write(aes.IV, 0, aes.IV.Length);

                using (CryptoStream cryptoStream = new CryptoStream(fsOut, aes.CreateEncryptor(), CryptoStreamMode.Write))
                using (FileStream fsIn = new FileStream(tempZip, FileMode.Open, FileAccess.Read))
                {
                    fsIn.CopyTo(cryptoStream);
                }
            }
        }

        // Clean up temporary file
        File.Delete(tempZip);
    }

    // Helper to generate cryptographically strong random bytes
    private static byte[] GenerateRandomBytes(int count)
    {
        byte[] bytes = new byte[count];
        using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(bytes);
        }
        return bytes;
    }
}
