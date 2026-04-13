using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractImagesFromTables
{
    public static void Main()
    {
        // Paths for temporary files
        const string imagePath = "sample.png";
        const string docPath = "sample.docx";
        const string zipPath = "extracted_images.zip";

        // Ensure previous files are removed
        if (File.Exists(imagePath)) File.Delete(imagePath);
        if (File.Exists(docPath)) File.Delete(docPath);
        if (File.Exists(zipPath)) File.Delete(zipPath);

        // -------------------------------------------------
        // Create a deterministic sample image (100x100 blue)
        // -------------------------------------------------
        Bitmap bitmap = new Bitmap(100, 100);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Blue);
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // Create a DOCX document with a table containing the image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table
        builder.StartTable();

        // First cell: insert the sample image
        builder.InsertCell();
        builder.InsertImage(imagePath);

        // Second cell: some text (optional)
        builder.InsertCell();
        builder.Writeln("Sample text");

        // End the row and the table
        builder.EndRow();
        builder.EndTable();

        // Save the document
        doc.Save(docPath);

        // -------------------------------------------------
        // Load the document for extraction
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Prepare a zip archive to store extracted images
        using (FileStream zipFileStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Create))
        {
            int imageIndex = 0;

            // Iterate through all tables in the document
            NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables.OfType<Table>())
            {
                // Get all Shape nodes inside the current table
                NodeCollection shapes = table.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapes.OfType<Shape>())
                {
                    if (shape.HasImage)
                    {
                        // Determine file extension based on image type
                        string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                        string entryName = $"image{imageIndex}{extension}";

                        // Save image to a memory stream
                        using (MemoryStream imageStream = new MemoryStream())
                        {
                            shape.ImageData.Save(imageStream);
                            imageStream.Position = 0; // Reset before reading

                            // Add entry to zip archive
                            ZipArchiveEntry entry = zipArchive.CreateEntry(entryName);
                            using (Stream entryStream = entry.Open())
                            {
                                imageStream.CopyTo(entryStream);
                            }
                        }

                        imageIndex++;
                    }
                }
            }

            // Validate that at least one image was extracted
            if (imageIndex == 0)
                throw new InvalidOperationException("No images were extracted from tables.");

            // Ensure the zip archive is flushed and closed by disposing the using blocks
        }

        // Final validation: the zip file must exist and contain entries
        if (!File.Exists(zipPath))
            throw new FileNotFoundException("The zip archive was not created.");

        // Optional: verify that the zip contains files
        using (FileStream zipFileStream = new FileStream(zipPath, FileMode.Open))
        using (ZipArchive zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Read))
        {
            if (!zipArchive.Entries.Any())
                throw new InvalidOperationException("The zip archive is empty.");
        }

        // Cleanup temporary files (optional)
        // File.Delete(imagePath);
        // File.Delete(docPath);
    }
}
