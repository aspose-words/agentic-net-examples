using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string imagesDir = Path.Combine(artifactsDir, "Extracted");
        Directory.CreateDirectory(imagesDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (input.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "input.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            // Simple deterministic drawing – a filled rectangle.
            graphics.FillRectangle(new SolidBrush(Color.DarkBlue), 10, 10, 80, 80);
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCX document that contains a table with images.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a 2x2 table and insert the sample image into each cell.
        builder.StartTable();
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                // Insert the image inline.
                builder.InsertImage(sampleImagePath);
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images that reside inside tables.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        int imageIndex = 0;

        // Prepare the zip archive that will hold the extracted images.
        string zipPath = Path.Combine(artifactsDir, "ExtractedImages.zip");
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            // Iterate over all tables in the document.
            NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables.OfType<Table>())
            {
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
                                // Determine file extension based on image type.
                                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                                string imageFileName = $"image_{imageIndex}{extension}";
                                string imageFilePath = Path.Combine(imagesDir, imageFileName);

                                // Save the image to a temporary file.
                                shape.ImageData.Save(imageFilePath);

                                // Add the image file to the zip archive.
                                ZipArchiveEntry entry = archive.CreateEntry(imageFileName);
                                using (FileStream imgStream = new FileStream(imageFilePath, FileMode.Open, FileAccess.Read))
                                using (Stream entryStream = entry.Open())
                                {
                                    imgStream.CopyTo(entryStream);
                                }

                                imageIndex++;
                            }
                        }
                    }
                }
            }

            // Validation: ensure at least one image was extracted.
            if (imageIndex == 0)
                throw new InvalidOperationException("No images were extracted from tables.");
        }

        // Optional cleanup of temporary extracted images.
        if (Directory.Exists(imagesDir))
            Directory.Delete(imagesDir, true);
    }
}
