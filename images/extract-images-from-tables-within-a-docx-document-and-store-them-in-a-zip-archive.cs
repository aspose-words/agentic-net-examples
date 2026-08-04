using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare working directories
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string imagePath = Path.Combine(workDir, "sample.png");
        string docPath = Path.Combine(workDir, "sample.docx");
        string extractDir = Path.Combine(workDir, "ExtractedImages");
        Directory.CreateDirectory(extractDir);
        string zipPath = Path.Combine(workDir, "ImagesFromTables.zip");

        // -------------------------------------------------
        // 1. Create a deterministic sample image (100x100 white)
        // -------------------------------------------------
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Draw a simple rectangle to make the image non‑empty
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
            {
                graphics.DrawRectangle(pen, 10, 10, 80, 80);
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Build a DOCX containing a table with images
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 2x2 table; put the image in the first cell
        builder.StartTable();

        // Cell 1,1 – contains the image
        builder.InsertCell();
        builder.InsertImage(imagePath);

        // Cell 1,2 – plain text
        builder.InsertCell();
        builder.Writeln("No image here");

        // End first row
        builder.EndRow();

        // Second row – both cells plain text
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Save the document
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images that reside inside tables
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);

        int imageIndex = 0;
        foreach (Table table in tables.OfType<Table>())
        {
            // Find all Shape nodes inside the current table
            NodeCollection shapes = table.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string extractedFile = Path.Combine(extractDir, $"image_{imageIndex}{extension}");
                    shape.ImageData.Save(extractedFile);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from tables.");

        // -------------------------------------------------
        // 4. Pack the extracted images into a ZIP archive
        // -------------------------------------------------
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (string file in Directory.GetFiles(extractDir))
            {
                string entryName = Path.GetFileName(file);
                archive.CreateEntryFromFile(file, entryName);
            }
        }

        // Optional cleanup of temporary extracted images
        foreach (string file in Directory.GetFiles(extractDir))
            File.Delete(file);
        Directory.Delete(extractDir);

        // Indicate successful completion (no interactive prompts)
        Console.WriteLine($"Extracted {imageIndex} image(s) and created archive: {zipPath}");
    }
}
