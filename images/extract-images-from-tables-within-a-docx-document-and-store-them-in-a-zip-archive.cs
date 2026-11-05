using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractImagesFromTables
{
    public static void Main()
    {
        // Prepare deterministic file names and folders
        string workingDir = Directory.GetCurrentDirectory();
        string sampleImagePath = Path.Combine(workingDir, "sample.png");
        string docPath = Path.Combine(workingDir, "sample.docx");
        string extractedFolder = Path.Combine(workingDir, "ExtractedImages");
        string zipPath = Path.Combine(workingDir, "ImagesFromTables.zip");

        // Ensure clean state
        if (File.Exists(sampleImagePath)) File.Delete(sampleImagePath);
        if (File.Exists(docPath)) File.Delete(docPath);
        if (Directory.Exists(extractedFolder)) Directory.Delete(extractedFolder, true);
        if (File.Exists(zipPath)) File.Delete(zipPath);

        // -------------------------------------------------
        // Create a deterministic sample image (100x100 white with a blue rectangle)
        // -------------------------------------------------
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
        try
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }

            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }

        // -------------------------------------------------
        // Create a DOCX document with a table containing images
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 2x2 table
        Table table = builder.StartTable();

        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();                     // Start a new cell
                builder.InsertImage(sampleImagePath);     // Insert the sample image into the cell
                // No explicit EndCell call needed; InsertCell moves to the next cell.
            }
            builder.EndRow(); // End the current row
        }

        builder.EndTable(); // End the table

        // Save the document
        doc.Save(docPath);

        // -------------------------------------------------
        // Load the document and extract images from tables only
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);

        Directory.CreateDirectory(extractedFolder);
        int imageIndex = 0;

        foreach (Table tbl in tables)
        {
            // Find all Shape nodes inside the current table
            NodeCollection shapes = tbl.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFile = Path.Combine(extractedFolder, $"image-{imageIndex}.png");
                    shape.ImageData.Save(imageFile);
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from tables.");

        // -------------------------------------------------
        // Create a zip archive containing the extracted images
        // -------------------------------------------------
        ZipFile.CreateFromDirectory(extractedFolder, zipPath);

        // Validate zip creation
        if (!File.Exists(zipPath))
            throw new InvalidOperationException("Failed to create the zip archive.");

        // Optional: Verify that the zip contains entries
        using (ZipArchive archive = ZipFile.OpenRead(zipPath))
        {
            if (archive.Entries.Count == 0)
                throw new InvalidOperationException("The zip archive is empty.");
        }

        // Cleanup temporary files (optional)
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
        // Directory.Delete(extractedFolder, true);
    }
}
