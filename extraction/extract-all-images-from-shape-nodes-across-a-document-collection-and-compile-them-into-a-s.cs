using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Base temporary directory for the example.
        string baseDir = Path.Combine(Path.GetTempPath(), "AsposeImagesExtraction");
        string docsDir = Path.Combine(baseDir, "Docs");
        string imagesDir = Path.Combine(baseDir, "Extracted");
        string zipPath = Path.Combine(baseDir, "AllImages.zip");

        // Ensure a clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(imagesDir);

        // Sample PNG image (1x1 pixel, transparent) encoded in Base64.
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        // Create a few sample documents, each containing an image.
        for (int docIndex = 0; docIndex < 3; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            using (MemoryStream imgStream = new MemoryStream(pngBytes))
            {
                builder.InsertImage(imgStream);
            }
            string docPath = Path.Combine(docsDir, $"Document{docIndex}.docx");
            doc.Save(docPath);
        }

        // List to keep track of extracted image file paths.
        List<string> extractedImageFiles = new List<string>();

        // Process each document in the collection.
        string[] docFiles = Directory.GetFiles(docsDir, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document loadedDoc = new Document(docFile);
            NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}{extension}";
                    string imagePath = Path.Combine(imagesDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImageFiles.Add(imagePath);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImageFiles.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document collection.");

        // Create a ZIP archive containing all extracted images.
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
        {
            foreach (string imageFile in extractedImageFiles)
            {
                string entryName = Path.GetFileName(imageFile);
                archive.CreateEntryFromFile(imageFile, entryName);
            }
        }

        // Verify that the ZIP file was created successfully.
        if (!File.Exists(zipPath) || new FileInfo(zipPath).Length == 0)
            throw new InvalidOperationException("Failed to create the ZIP archive with extracted images.");

        // Example completed successfully.
        Console.WriteLine("Images extracted and packaged into ZIP file at:");
        Console.WriteLine(zipPath);
    }
}
