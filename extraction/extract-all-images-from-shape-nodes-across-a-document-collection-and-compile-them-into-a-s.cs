using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder to hold temporary documents.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeImageExtraction");
        if (Directory.Exists(workFolder))
            Directory.Delete(workFolder, true);
        Directory.CreateDirectory(workFolder);

        // Base64 encoded 1x1 pixel PNG image.
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        // Create a few sample documents each containing images.
        int documentCount = 2;
        for (int docIndex = 0; docIndex < documentCount; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert two images into each document.
            for (int img = 0; img < 2; img++)
            {
                using (MemoryStream imgStream = new MemoryStream(pngBytes))
                {
                    builder.InsertImage(imgStream);
                }
                builder.Writeln(); // Separate images with a paragraph.
            }

            string docPath = Path.Combine(workFolder, $"SampleDoc{docIndex}.docx");
            doc.Save(docPath);
        }

        // Path for the resulting ZIP archive.
        string zipPath = Path.Combine(workFolder, "ExtractedImages.zip");

        // Create the ZIP archive and add all extracted images.
        using (FileStream zipFile = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipFile, ZipArchiveMode.Create))
        {
            // Process each document in the work folder.
            string[] docFiles = Directory.GetFiles(workFolder, "*.docx");
            foreach (string docFile in docFiles)
            {
                Document loadedDoc = new Document(docFile);
                var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                      .OfType<Shape>()
                                      .Where(s => s.HasImage);

                int imageIndex = 0;
                foreach (Shape shape in shapes)
                {
                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string entryName = $"image_doc{Path.GetFileNameWithoutExtension(docFile)}_{imageIndex}{extension}";

                    using (MemoryStream imgStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imgStream);
                        imgStream.Position = 0;

                        ZipArchiveEntry entry = archive.CreateEntry(entryName);
                        using (Stream entryStream = entry.Open())
                        {
                            imgStream.CopyTo(entryStream);
                        }
                    }

                    imageIndex++;
                }
            }
        }

        // Validation: ensure the ZIP file was created and contains entries.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("ZIP archive was not created.");

        using (FileStream zipCheck = new FileStream(zipPath, FileMode.Open))
        using (ZipArchive archive = new ZipArchive(zipCheck, ZipArchiveMode.Read))
        {
            if (archive.Entries.Count == 0)
                throw new InvalidOperationException("No images were added to the ZIP archive.");
        }

        // Cleanup temporary files (optional).
        // Directory.Delete(workFolder, true);
    }
}
