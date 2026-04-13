using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    // Password used for encrypting the final ZIP archive
    private const string ZipPassword = "Secret123";

    public static void Main()
    {
        // Define working folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        // Prepare folders
        PrepareDirectory(inputDocsDir);
        PrepareDirectory(imagesDir);
        PrepareDirectory(outputDir);

        // 1. Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Create sample DOCX files that contain the image
        CreateSampleDocument(Path.Combine(inputDocsDir, "Sample1.docx"), sampleImagePath);
        CreateSampleDocument(Path.Combine(inputDocsDir, "Sample2.docx"), sampleImagePath);

        // 3. Batch process each DOC/DOCX file: extract images
        int totalExtracted = 0;
        foreach (string docPath in Directory.GetFiles(inputDocsDir, "*.*", SearchOption.TopDirectoryOnly)
                                            .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                                        f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
        {
            totalExtracted += ExtractImagesFromDocument(docPath, imagesDir);
        }

        // Validate that at least one image was extracted
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // 4. Create a ZIP archive from the extracted images
        string zipPath = Path.Combine(outputDir, "Images.zip");
        // Resolve ambiguity by using the System.IO.Compression enum explicitly
        ZipFile.CreateFromDirectory(imagesDir, zipPath, System.IO.Compression.CompressionLevel.Optimal, false);

        // 5. Encrypt the ZIP archive with a password (AES‑CBC)
        string encryptedZipPath = Path.Combine(outputDir, "ImagesProtected.zip");
        EncryptFileWithPassword(zipPath, encryptedZipPath, ZipPassword);

        // Validate that the encrypted ZIP exists
        if (!File.Exists(encryptedZipPath))
            throw new FileNotFoundException("Failed to create the encrypted ZIP archive.");

        // Cleanup intermediate ZIP (optional)
        File.Delete(zipPath);
    }

    // Ensures a clean directory exists
    private static void PrepareDirectory(string path)
    {
        if (Directory.Exists(path))
        {
            foreach (string file in Directory.GetFiles(path))
                File.Delete(file);
            foreach (string dir in Directory.GetDirectories(path))
                Directory.Delete(dir, true);
        }
        else
        {
            Directory.CreateDirectory(path);
        }
    }

    // Creates a deterministic PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Use explicit Aspose.Drawing types as required by the rules
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.FromArgb(255, 100, 150, 200)))
        {
            g.FillRectangle(brush, 20, 20, width - 40, height - 40);
        }
        // Save as PNG
        bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        g.Dispose();
        bitmap.Dispose();
    }

    // Creates a DOCX file with the provided image inserted
    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Extracts all images from a document and saves them to the target folder
    private static int ExtractImagesFromDocument(string docPath, string targetFolder)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image_{imageIndex}{extension}";
                string fullPath = Path.Combine(targetFolder, imageFileName);
                shape.ImageData.Save(fullPath);
                imageIndex++;
            }
        }

        return imageIndex;
    }

    // Encrypts a file using AES‑CBC with a password‑derived key
    private static void EncryptFileWithPassword(string inputFile, string outputFile, string password)
    {
        // Derive a 256‑bit key from the password using SHA‑256
        byte[] key;
        using (SHA256 sha = SHA256.Create())
        {
            key = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
        }

        // Generate a random IV
        byte[] iv = new byte[16];
        using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(iv);
        }

        using (FileStream fsInput = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
        using (FileStream fsOutput = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
        {
            // Write IV at the beginning of the file (needed for decryption)
            fsOutput.Write(iv, 0, iv.Length);

            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (CryptoStream cryptoStream = new CryptoStream(fsOutput, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    fsInput.CopyTo(cryptoStream);
                }
            }
        }
    }
}
