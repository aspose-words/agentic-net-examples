using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "WorkDir");
        Directory.CreateDirectory(workDir);

        // Base64 encoded 1x1 PNG image.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Create sample documents with images.
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= 2; i++)
        {
            string docPath = Path.Combine(workDir, $"doc{i}.docx");
            CreateSampleDocumentWithImage(docPath, pngBytes);
            docPaths.Add(docPath);
        }

        // Path for the resulting ZIP archive.
        string zipPath = Path.Combine(workDir, "extracted_images.zip");

        // Extract images from all documents and add them to the ZIP.
        using (FileStream zipFileStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Create))
        {
            foreach (string docPath in docPaths)
            {
                Document loaded = new Document(docPath);

                // Find all shape nodes that contain images.
                IEnumerable<Shape> imageShapes = loaded.GetChildNodes(NodeType.Shape, true)
                    .OfType<Shape>()
                    .Where(s => s.HasImage);

                int imageIndex = 0;
                foreach (Shape shape in imageShapes)
                {
                    // Determine a deterministic entry name.
                    string extension = shape.ImageData.ImageType.ToString().ToLower(); // Use ImageType instead of ImageFormat
                    string entryName = $"{Path.GetFileNameWithoutExtension(docPath)}_image_{imageIndex}.{extension}";
                    ZipArchiveEntry entry = zipArchive.CreateEntry(entryName, CompressionLevel.Optimal);
                    using (Stream entryStream = entry.Open())
                    {
                        byte[] imageBytes = shape.ImageData.ImageBytes;
                        entryStream.Write(imageBytes, 0, imageBytes.Length);
                    }
                    imageIndex++;
                }
            }
        }

        // Validation: ensure the ZIP file was created and contains at least one entry.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("The ZIP archive was not created.");

        using (FileStream zipReadStream = new FileStream(zipPath, FileMode.Open))
        using (ZipArchive zipRead = new ZipArchive(zipReadStream, ZipArchiveMode.Read))
        {
            if (zipRead.Entries.Count == 0)
                throw new InvalidOperationException("No images were extracted into the ZIP archive.");
        }

        // Cleanup: optional removal of temporary files can be added here.
    }

    private static void CreateSampleDocumentWithImage(string filePath, byte[] imageBytes)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            builder.InsertImage(imageStream);
        }
        doc.Save(filePath);
    }
}
