using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for sample documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
        Directory.CreateDirectory(docsFolder);

        // Create sample documents each containing an image shape.
        CreateSampleDocument(Path.Combine(docsFolder, "doc1.docx"));
        CreateSampleDocument(Path.Combine(docsFolder, "doc2.docx"));

        // Path for the resulting ZIP archive.
        string zipPath = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages.zip");

        // Ensure any existing ZIP is removed.
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        // Create the ZIP archive and add extracted images.
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            // Process each document in the collection.
            foreach (string docPath in Directory.GetFiles(docsFolder, "*.docx"))
            {
                Document doc = new Document(docPath);

                // Find all shape nodes that contain images.
                List<Shape> imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                                            .OfType<Shape>()
                                            .Where(s => s.HasImage)
                                            .ToList();

                // Extract each image and add it to the ZIP.
                for (int i = 0; i < imageShapes.Count; i++)
                {
                    Shape shape = imageShapes[i];
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imageStream);
                        imageStream.Position = 0;

                        string imageExtension = GetImageExtension(shape.ImageData.ImageType);
                        string entryName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{i + 1}{imageExtension}";
                        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                        using (Stream entryStream = entry.Open())
                        {
                            imageStream.CopyTo(entryStream);
                        }
                    }
                }
            }
        }

        // Validation: ensure the ZIP file was created and contains entries.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("ZIP archive was not created.");

        using (ZipArchive archive = ZipFile.OpenRead(zipPath))
        {
            if (archive.Entries.Count == 0)
                throw new InvalidOperationException("No images were added to the ZIP archive.");
        }
    }

    // Helper method to create a sample document with a single PNG image.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A tiny 1x1 pixel PNG image (base64 encoded).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imageStream);
        }

        doc.Save(filePath);
    }

    // Helper to map Aspose.Words.Drawing.ImageType to a file extension.
    private static string GetImageExtension(ImageType imageType)
    {
        return imageType switch
        {
            ImageType.Jpeg => ".jpg",
            ImageType.Png => ".png",
            ImageType.Gif => ".gif",
            ImageType.Bmp => ".bmp",
            ImageType.Emf => ".emf",
            ImageType.Wmf => ".wmf",
            _ => ".img"
        };
    }
}
