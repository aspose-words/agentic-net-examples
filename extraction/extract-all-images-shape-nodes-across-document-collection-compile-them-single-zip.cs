using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

// Alias to avoid ambiguity with Aspose.Words.Saving.CompressionLevel
using SystemCompressionLevel = System.IO.Compression.CompressionLevel;

public class ShapeImageExtractor
{
    /// <summary>
    /// Extracts all images from shape nodes in the specified Word documents and stores them in a single ZIP archive.
    /// </summary>
    /// <param name="documentPaths">Array of file paths to the source Word documents.</param>
    /// <param name="zipFilePath">Destination path for the created ZIP archive.</param>
    public static void ExtractImagesToZip(string[] documentPaths, string zipFilePath)
    {
        // Ensure the output directory exists.
        string zipDirectory = Path.GetDirectoryName(zipFilePath) ?? string.Empty;
        if (!string.IsNullOrEmpty(zipDirectory) && !Directory.Exists(zipDirectory))
        {
            Directory.CreateDirectory(zipDirectory);
        }

        // Create (or overwrite) the ZIP archive.
        using (FileStream zipToOpen = new FileStream(zipFilePath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            int globalImageIndex = 0; // Unique index across all documents.

            foreach (string docPath in documentPaths)
            {
                // Load the document.
                Document doc = new Document(docPath);

                // Retrieve all Shape nodes (deep search).
                NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

                // Filter shapes that actually contain an image.
                IEnumerable<Shape> imageShapes = shapeNodes
                    .OfType<Shape>()
                    .Where(s => s.HasImage);

                foreach (Shape shape in imageShapes)
                {
                    // Determine file extension based on the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    // Build a unique file name for the entry inside the ZIP.
                    string entryName = $"Image_{globalImageIndex}{extension}";

                    // Save the image to a memory stream (no intermediate file needed).
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        shape.ImageData.Save(imageStream);
                        imageStream.Position = 0; // Reset for reading.

                        // Create a new entry in the ZIP archive and copy the image data.
                        ZipArchiveEntry entry = archive.CreateEntry(entryName, SystemCompressionLevel.Optimal);
                        using (Stream entryStream = entry.Open())
                        {
                            imageStream.CopyTo(entryStream);
                        }
                    }

                    globalImageIndex++;
                }
            }
        }
    }

    // Example usage.
    public static void Main()
    {
        // Create a temporary folder for the demo files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "ShapeImageExtractorDemo");
        Directory.CreateDirectory(tempFolder);

        // Paths to the source documents (generated on the fly).
        string[] docs = new string[2];
        for (int i = 0; i < docs.Length; i++)
        {
            string docPath = Path.Combine(tempFolder, $"Document{i + 1}.docx");
            // Create a simple document. No images are added, but the code works even if there are none.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i + 1}.");
            doc.Save(docPath);
            docs[i] = docPath;
        }

        // Destination ZIP file.
        string zipPath = Path.Combine(tempFolder, "AllImages.zip");

        // Perform extraction.
        ExtractImagesToZip(docs, zipPath);

        Console.WriteLine($"Extraction completed. ZIP file created at: {zipPath}");
    }
}
