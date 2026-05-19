using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;

public class ExtractImagesFromTables
{
    public static void Main()
    {
        // Define paths for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string imagePath = Path.Combine(artifactsDir, "sample.png");
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        string zipPath = Path.Combine(artifactsDir, "extracted_images.zip");

        // Ensure the output directory exists.
        Directory.CreateDirectory(artifactsDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Draw a simple rectangle for visual distinction.
        graphics.FillRectangle(new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Blue), 20, 20, imgWidth - 40, imgHeight - 40);
        // Save the image to a local file.
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a DOCX document with a table that contains the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with 2 rows and 2 columns.
        builder.StartTable();

        // First row, first cell: insert the image.
        builder.InsertCell();
        builder.InsertImage(imagePath);
        // First row, second cell: some text.
        builder.InsertCell();
        builder.Writeln("Cell with text");

        // End first row.
        builder.EndRow();

        // Second row, first cell: text.
        builder.InsertCell();
        builder.Writeln("Another cell");
        // Second row, second cell: text.
        builder.InsertCell();
        builder.Writeln("Yet another cell");

        // End second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images that reside inside tables.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        List<(byte[] Data, string Extension)> extractedImages = new List<(byte[], string)>();

        // Get all tables in the document.
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables.OfType<Table>())
        {
            // Iterate through each cell of the table.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    // Find all Shape nodes inside the cell.
                    NodeCollection shapes = cell.GetChildNodes(NodeType.Shape, true);
                    foreach (Shape shape in shapes.OfType<Shape>())
                    {
                        if (shape.HasImage)
                        {
                            // Save image data to a memory stream.
                            using (MemoryStream ms = new MemoryStream())
                            {
                                shape.ImageData.Save(ms);
                                ms.Position = 0;
                                byte[] imageBytes = ms.ToArray();
                                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                                extractedImages.Add((imageBytes, extension));
                            }
                        }
                    }
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were found inside tables.");

        // -------------------------------------------------
        // 4. Store the extracted images into a ZIP archive.
        // -------------------------------------------------
        // Delete existing zip if present.
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            for (int i = 0; i < extractedImages.Count; i++)
            {
                string entryName = $"image_{i}{extractedImages[i].Extension}";
                ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                using (Stream entryStream = entry.Open())
                using (MemoryStream imageStream = new MemoryStream(extractedImages[i].Data))
                {
                    imageStream.CopyTo(entryStream);
                }
            }
        }

        // -------------------------------------------------
        // 5. Completion message.
        // -------------------------------------------------
        Console.WriteLine($"Extracted {extractedImages.Count} image(s) from tables and saved to '{zipPath}'.");
    }
}
