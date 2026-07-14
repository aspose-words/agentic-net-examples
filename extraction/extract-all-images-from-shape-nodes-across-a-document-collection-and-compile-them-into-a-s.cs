using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directories for the sample documents and extracted images.
        string baseDir = Directory.GetCurrentDirectory();
        string docsDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string zipPath = Path.Combine(baseDir, "AllImages.zip");

        // Ensure clean start.
        if (Directory.Exists(docsDir)) Directory.Delete(docsDir, true);
        if (Directory.Exists(imagesDir)) Directory.Delete(imagesDir, true);
        if (File.Exists(zipPath)) File.Delete(zipPath);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(imagesDir);

        // Base64 encoded 1x1 pixel PNG (used for all sample images).
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        // Create a few sample documents, each containing one image.
        for (int docIndex = 0; docIndex < 3; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            using (MemoryStream imgStream = new MemoryStream(pngBytes))
            {
                builder.InsertImage(imgStream);
            }
            string docPath = Path.Combine(docsDir, $"SampleDoc{docIndex}.docx");
            doc.Save(docPath);
        }

        // Extract images from all documents in the collection.
        int globalImageIndex = 0;
        foreach (string docPath in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document loaded = new Document(docPath);
            NodeCollection shapes = loaded.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"Img_{globalImageIndex}{extension}";
                    string imageFullPath = Path.Combine(imagesDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    globalImageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (globalImageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document collection.");

        // Create a ZIP archive containing all extracted images.
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (string imagePath in Directory.GetFiles(imagesDir))
            {
                string entryName = Path.GetFileName(imagePath);
                archive.CreateEntryFromFile(imagePath, entryName);
            }
        }

        // Final validation: ensure the ZIP file exists and contains entries.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("ZIP archive was not created.");

        using (FileStream zipRead = new FileStream(zipPath, FileMode.Open))
        using (ZipArchive archive = new ZipArchive(zipRead, ZipArchiveMode.Read))
        {
            if (archive.Entries.Count == 0)
                throw new InvalidOperationException("ZIP archive is empty.");
        }

        // The example runs to completion without interactive prompts.
    }
}
