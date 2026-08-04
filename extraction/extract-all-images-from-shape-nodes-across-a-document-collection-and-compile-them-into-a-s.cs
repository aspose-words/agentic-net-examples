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
        // Prepare a temporary folder for sample documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        if (Directory.Exists(docsFolder))
            Directory.Delete(docsFolder, true);
        Directory.CreateDirectory(docsFolder);

        // Base64-encoded 1x1 pixel PNG image.
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        // Create a few sample documents, each containing the same image.
        int documentCount = 3;
        List<string> docPaths = new List<string>();
        for (int i = 0; i < documentCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            using (MemoryStream imgStream = new MemoryStream(pngBytes))
            {
                builder.InsertImage(imgStream);
            }
            string docPath = Path.Combine(docsFolder, $"SampleDoc{i}.docx");
            doc.Save(docPath);
            docPaths.Add(docPath);
        }

        // Path for the resulting ZIP archive.
        string zipPath = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages.zip");
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        // Create the ZIP archive and add all extracted images.
        using (FileStream zipFileStream = new FileStream(zipPath, FileMode.CreateNew))
        using (ZipArchive zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Create))
        {
            for (int docIndex = 0; docIndex < docPaths.Count; docIndex++)
            {
                string path = docPaths[docIndex];
                Document loadedDoc = new Document(path);

                // Retrieve all shape nodes in the document.
                IEnumerable<Shape> shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                                    .OfType<Shape>()
                                                    .Where(s => s.HasImage);

                int imageIndex = 0;
                foreach (Shape shape in shapes)
                {
                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string entryName = $"Doc{docIndex}_Image{imageIndex}{extension}";

                    // Create a new entry in the ZIP archive.
                    ZipArchiveEntry entry = zipArchive.CreateEntry(entryName);
                    using (Stream entryStream = entry.Open())
                    {
                        // Save the image data directly into the ZIP entry stream.
                        shape.ImageData.Save(entryStream);
                    }
                    imageIndex++;
                }

                // Validate that at least one image was found in the current document.
                if (!shapes.Any())
                    throw new InvalidOperationException($"No images found in document: {path}");
            }
        }

        // Verify that the ZIP file was created and contains entries.
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("The ZIP archive was not created.");

        using (FileStream zipReadStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read))
        using (ZipArchive zipRead = new ZipArchive(zipReadStream, ZipArchiveMode.Read))
        {
            if (zipRead.Entries.Count == 0)
                throw new InvalidOperationException("The ZIP archive contains no entries.");
        }

        // Cleanup temporary documents (optional).
        Directory.Delete(docsFolder, true);
    }
}
