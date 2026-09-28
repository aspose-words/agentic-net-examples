using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string sampleImagePath = "sample.png";
        const string docPath = "sample.docx";
        const string zipPath = "TableImages.zip";

        // 1. Create a deterministic sample image using Aspose.Drawing
        const int imgWidth = 100;
        const int imgHeight = 100;
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // (Optional) draw something simple here if desired
        graphics.Dispose();
        bitmap.Save(sampleImagePath);
        bitmap.Dispose();

        // 2. Create a new document with a table containing the sample image in each cell
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 2x2 table
        builder.StartTable();
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                // Start a new cell
                builder.InsertCell();

                // Insert the image into the current cell
                builder.InsertImage(sampleImagePath);
            }
            // End the current row
            builder.EndRow();
        }
        // End the table
        builder.EndTable();

        // Save the document to disk
        doc.Save(docPath);

        // 3. Load the document (simulating a separate load step)
        Document loadedDoc = new Document(docPath);

        // 4. Extract images that are inside tables only
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);
        int imageIndex = 0;

        // Prepare the zip archive for output
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
        {
            foreach (Table tbl in tables)
            {
                // Find all Shape nodes within the current table
                NodeCollection shapes = tbl.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapes)
                {
                    if (shape.HasImage)
                    {
                        // Save the image to a memory stream
                        using (MemoryStream imgStream = new MemoryStream())
                        {
                            shape.ImageData.Save(imgStream);
                            imgStream.Position = 0; // Reset before reading

                            // Add the image to the zip archive with a deterministic name
                            string entryName = $"image-{imageIndex}.png";
                            ZipArchiveEntry entry = archive.CreateEntry(entryName);
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

        // 5. Validation: ensure at least one image was extracted
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were extracted from tables.");
        }

        // 6. Validation: ensure the zip file exists
        if (!File.Exists(zipPath))
        {
            throw new FileNotFoundException("The zip archive was not created.", zipPath);
        }

        // Cleanup temporary files (optional)
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
