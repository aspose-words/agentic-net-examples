using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ExtractImagesFromTables
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string inputDocPath = @"C:\Docs\InputDocument.docx";

        // Path where the resulting ZIP archive will be saved.
        const string outputZipPath = @"C:\Docs\ExtractedImages.zip";

        // Load the Word document using the Aspose.Words Document constructor (lifecycle rule).
        Document doc = new Document(inputDocPath);

        // Create a new ZIP archive (System.IO.Compression) to hold the extracted images.
        using (FileStream zipToCreate = new FileStream(outputZipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToCreate, ZipArchiveMode.Create))
        {
            int imageIndex = 0;

            // Retrieve all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

            foreach (Table table in tables.OfType<Table>())
            {
                // Within each table, find all Shape nodes (these can contain images).
                NodeCollection shapes = table.GetChildNodes(NodeType.Shape, true);

                foreach (Shape shape in shapes.OfType<Shape>())
                {
                    if (shape.HasImage)
                    {
                        // Determine a suitable file extension for the image based on its type.
                        string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                        // Build a unique file name for the image inside the ZIP.
                        string entryName = $"Image_{imageIndex}{extension}";

                        // Create a new entry in the ZIP archive.
                        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);

                        // Write the image bytes into the ZIP entry.
                        using (Stream entryStream = entry.Open())
                        using (MemoryStream imageStream = new MemoryStream())
                        {
                            // Save the image data to a memory stream.
                            shape.ImageData.Save(imageStream);
                            imageStream.Position = 0;

                            // Copy the image bytes to the ZIP entry stream.
                            imageStream.CopyTo(entryStream);
                        }

                        imageIndex++;
                    }
                }
            }
        }

        // At this point the ZIP file contains all images that were located inside tables.
        Console.WriteLine($"Extraction complete. {Path.GetFileName(outputZipPath)} created.");
    }
}
