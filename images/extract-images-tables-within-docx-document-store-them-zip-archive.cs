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
        // Use paths relative to the current directory.
        string inputDocPath = Path.Combine(Environment.CurrentDirectory, "InputDocument.docx");
        string outputZipPath = Path.Combine(Environment.CurrentDirectory, "ExtractedImages.zip");

        // Ensure a document exists. If not, create a simple one with a table containing an image.
        if (!File.Exists(inputDocPath))
        {
            CreateSampleDocumentWithImage(inputDocPath);
        }

        // Load the Word document.
        Document doc = new Document(inputDocPath);

        // Create a new ZIP archive to hold the extracted images.
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
                            shape.ImageData.Save(imageStream);
                            imageStream.Position = 0;
                            imageStream.CopyTo(entryStream);
                        }

                        imageIndex++;
                    }
                }
            }
        }

        Console.WriteLine($"Extraction complete. {Path.GetFileName(outputZipPath)} created.");
    }

    private static void CreateSampleDocumentWithImage(string path)
    {
        // A tiny 1x1 pixel PNG (transparent) encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with one cell.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell with image:");
        builder.InsertCell();

        // Insert the image into the second cell.
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            builder.InsertImage(ms);
        }

        builder.EndRow();
        builder.EndTable();

        doc.Save(path);
    }
}
