using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;          // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare folders.
        // -----------------------------------------------------------------
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imagesDir);

        // -----------------------------------------------------------------
        // 2. Create a deterministic sample image (sample.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 3. Build a DOCX document that contains a 2x2 table with images.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                // Begin a new cell.
                builder.InsertCell();

                // Insert the sample image into the cell.
                builder.InsertImage(sampleImagePath);
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "TableWithImages.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 4. Load the document and extract images that reside inside tables.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);

        int imageIndex = 0;
        foreach (Table tbl in tables)
        {
            foreach (Row row in tbl.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    // Find all Shape nodes inside the cell.
                    NodeCollection shapes = cell.GetChildNodes(NodeType.Shape, true);
                    foreach (Shape shape in shapes)
                    {
                        if (shape.HasImage)
                        {
                            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                            string imageFileName = $"image_{imageIndex}{extension}";
                            string imageFullPath = Path.Combine(imagesDir, imageFileName);
                            shape.ImageData.Save(imageFullPath);
                            imageIndex++;
                        }
                    }
                }
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the tables.");

        // -----------------------------------------------------------------
        // 5. Pack the extracted images into a ZIP archive.
        // -----------------------------------------------------------------
        string zipPath = Path.Combine(artifactsDir, "ExtractedImages.zip");
        using (FileStream zipStream = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            foreach (string filePath in Directory.GetFiles(imagesDir))
            {
                string entryName = Path.GetFileName(filePath);
                archive.CreateEntryFromFile(filePath, entryName);
            }
        }

        // All files are written to the "Artifacts" folder.
    }
}
