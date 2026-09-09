using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Drawing;   // Aspose.Drawing provides Bitmap, Graphics, Color, etc.

// Alias the System compression level to avoid ambiguity with Aspose.Words.Saving.CompressionLevel
using SystemCompressionLevel = System.IO.Compression.CompressionLevel;

public class Program
{
    // Password used to protect the final ZIP archive.
    private const string ZipPassword = "SecretPassword";

    public static void Main()
    {
        // Prepare folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "DemoArtifacts");
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // Step 1: Create deterministic sample images.
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImage1, 200, 150, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(sampleImage2, 150, 200, Aspose.Drawing.Color.LightCoral);

        // Step 2: Create sample DOCX files that contain the images.
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc1.docx"), sampleImage1);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Doc2.docx"), sampleImage2);

        // Step 3: Batch process each DOCX, extract all images.
        int globalImageIndex = 0;
        foreach (string docPath in Directory.GetFiles(inputDocsDir, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_img{globalImageIndex}{extension}";
                    string imageFullPath = Path.Combine(imagesDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    globalImageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (Directory.GetFiles(imagesDir).Length == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Step 4: Create a ZIP archive from the extracted images.
        string zipPath = Path.Combine(outputDir, "Images.zip");
        ZipFile.CreateFromDirectory(imagesDir, zipPath, SystemCompressionLevel.Optimal, false);

        // Step 5: Encrypt the ZIP archive with a password (AES encryption).
        string protectedZipPath = Path.Combine(outputDir, "Images_protected.zip");
        EncryptFile(zipPath, protectedZipPath, ZipPassword);

        // Clean up the unencrypted ZIP.
        File.Delete(zipPath);

        // Final validation.
        if (!File.Exists(protectedZipPath) || new FileInfo(protectedZipPath).Length == 0)
            throw new InvalidOperationException("Failed to create the password‑protected ZIP archive.");

        Console.WriteLine("Images extracted and password‑protected ZIP created at:");
        Console.WriteLine(protectedZipPath);
    }

    // Creates a simple bitmap image using Aspose.Drawing and saves it to the specified path.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(backgroundColor);
            // Additional deterministic drawing can be added here if needed.
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX document that contains a single image.
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document containing an image:");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Encrypts the source file using AES and writes the encrypted data to the destination file.
    private static void EncryptFile(string sourcePath, string destinationPath, string password)
    {
        // Generate a random 16‑byte salt.
        byte[] salt = new byte[16];
        RandomNumberGenerator.Fill(salt);

        // Derive a 256‑bit key and a 128‑bit IV from the password using SHA‑256.
        using (var pdb = new Rfc2898DeriveBytes(password, salt, 100_000, HashAlgorithmName.SHA256))
        {
            byte[] key = pdb.GetBytes(32); // 256 bits
            byte[] iv = pdb.GetBytes(16);  // 128 bits

            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (FileStream fsInput = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
                using (FileStream fsOutput = new FileStream(destinationPath, FileMode.Create, FileAccess.Write))
                {
                    // Write the salt at the beginning so it can be used for decryption.
                    fsOutput.Write(salt, 0, salt.Length);

                    using (CryptoStream cryptoStream = new CryptoStream(fsOutput, aes.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        fsInput.CopyTo(cryptoStream);
                    }
                }
            }
        }
    }
}
