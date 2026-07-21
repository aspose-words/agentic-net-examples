using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color

public class Program
{
    // Password used for encrypting the final ZIP archive
    private const string ZipPassword = "Secret123";

    public static void Main()
    {
        // Prepare working directories
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string docsDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(imagesDir);

        // Step 1: Create sample images and DOC files
        CreateSampleDocs(docsDir);

        // Step 2: Extract images from each DOC file
        List<string> extractedImageFiles = ExtractImagesFromDocs(docsDir, imagesDir);

        // Validate that at least one image was extracted
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Step 3: Create a ZIP archive containing the extracted images
        string zipPath = Path.Combine(baseDir, "ImagesArchive.zip");
        CreateZipArchive(extractedImageFiles, zipPath);

        // Step 4: Apply simple password‑based AES encryption to the ZIP file
        string protectedZipPath = Path.Combine(baseDir, "ImagesArchive_Protected.zip");
        EncryptFileWithPassword(zipPath, protectedZipPath, ZipPassword);

        // Validate final output
        if (!File.Exists(protectedZipPath))
            throw new FileNotFoundException("Protected ZIP archive was not created.", protectedZipPath);

        // Clean up intermediate ZIP (optional)
        File.Delete(zipPath);
    }

    // Creates a few DOC files, each containing a deterministic sample image
    private static void CreateSampleDocs(string docsFolder)
    {
        // Create a deterministic sample image
        string sampleImagePath = Path.Combine(docsFolder, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // Create two documents that embed the same image
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(docsFolder, $"Document{i}.docx");
            doc.Save(docPath);
        }
    }

    // Generates a simple white PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Extracts all images from DOC files in the source folder into the target folder
    private static List<string> ExtractImagesFromDocs(string sourceDocsFolder, string targetImagesFolder)
    {
        List<string> extractedFiles = new List<string>();
        string[] docFiles = Directory.GetFiles(sourceDocsFolder, "*.docx", SearchOption.TopDirectoryOnly);
        int imageCounter = 0;

        foreach (string docPath in docFiles)
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"Image_{imageCounter}{extension}";
                    string imageFullPath = Path.Combine(targetImagesFolder, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    extractedFiles.Add(imageFullPath);
                    imageCounter++;
                }
            }
        }

        return extractedFiles;
    }

    // Creates a ZIP archive from a list of files
    private static void CreateZipArchive(List<string> files, string zipFilePath)
    {
        using (FileStream zipToOpen = new FileStream(zipFilePath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            foreach (string filePath in files)
            {
                string entryName = Path.GetFileName(filePath);
                archive.CreateEntryFromFile(filePath, entryName);
            }
        }
    }

    // Encrypts a file using AES (CBC) with a password‑derived key
    private static void EncryptFileWithPassword(string inputPath, string outputPath, string password)
    {
        // Derive a 256‑bit key from the password using SHA‑256
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] key = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(password));

            // Generate a random IV and prepend it to the encrypted file
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.GenerateIV();

                using (FileStream inputFile = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                using (FileStream outputFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    // Write IV first
                    outputFile.Write(aes.IV, 0, aes.IV.Length);

                    using (CryptoStream cryptoStream = new CryptoStream(outputFile, aes.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        inputFile.CopyTo(cryptoStream);
                    }
                }
            }
        }
    }
}
