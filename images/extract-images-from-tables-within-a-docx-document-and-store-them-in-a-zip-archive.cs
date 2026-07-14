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
        // Define file paths
        string docPath = "sample.docx";
        string imagePath = "sampleImage.png";
        string extractedFolder = "extracted";
        string zipPath = "ExtractedImages.zip";

        // Ensure clean environment
        if (File.Exists(docPath)) File.Delete(docPath);
        if (File.Exists(imagePath)) File.Delete(imagePath);
        if (Directory.Exists(extractedFolder)) Directory.Delete(extractedFolder, true);
        if (File.Exists(zipPath)) File.Delete(zipPath);

        // 1. Create a sample image
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(imagePath);
        }

        // 2. Create a DOCX with a table containing images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 2x2 table
        builder.StartTable();

        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                builder.InsertImage(imagePath);
            }
            builder.EndRow();
        }

        builder.EndTable();

        doc.Save(docPath);

        // 3. Extract images from tables
        Document loadedDoc = new Document(docPath);
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);

        Directory.CreateDirectory(extractedFolder);
        int imageIndex = 1;

        foreach (Table tbl in tables)
        {
            foreach (Row row in tbl.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    NodeCollection shapes = cell.GetChildNodes(NodeType.Shape, true);
                    foreach (Shape shape in shapes)
                    {
                        if (shape.HasImage)
                        {
                            string outImagePath = Path.Combine(extractedFolder, $"image-{imageIndex}.png");
                            shape.ImageData.Save(outImagePath);
                            imageIndex++;
                        }
                    }
                }
            }
        }

        // Validate that at least one image was extracted
        string[] extractedFiles = Directory.GetFiles(extractedFolder);
        if (extractedFiles.Length == 0)
            throw new InvalidOperationException("No images were extracted from the tables.");

        // 4. Create a zip archive containing the extracted images
        using (ZipArchive zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
        {
            foreach (string filePath in extractedFiles)
            {
                zip.CreateEntryFromFile(filePath, Path.GetFileName(filePath));
            }
        }

        // Validate zip creation
        if (!File.Exists(zipPath) || new FileInfo(zipPath).Length == 0)
            throw new InvalidOperationException("Failed to create the zip archive.");

        // Program completed successfully
        Console.WriteLine("Image extraction and zipping completed.");
    }
}
