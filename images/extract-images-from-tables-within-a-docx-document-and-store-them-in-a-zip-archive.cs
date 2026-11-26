using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths
        const string sampleImagePath = "input.png";
        const string docPath = "sample.docx";
        const string outputZipPath = "ImagesFromTables.zip";
        const string tempImageFolder = "extracted_images";

        // Clean previous artifacts
        if (File.Exists(sampleImagePath)) File.Delete(sampleImagePath);
        if (File.Exists(docPath)) File.Delete(docPath);
        if (File.Exists(outputZipPath)) File.Delete(outputZipPath);
        if (Directory.Exists(tempImageFolder)) Directory.Delete(tempImageFolder, true);
        Directory.CreateDirectory(tempImageFolder);

        // 1. Create a deterministic sample image
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // 2. Create a DOCX with a 2x2 table containing the image in the first cell
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start table
        builder.StartTable();

        // Row 1, Cell 1 – image
        builder.InsertCell();
        builder.InsertImage(sampleImagePath);
        builder.EndRow();

        // Row 2, Cell 1 – text
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Row 3, Cell 1 – text
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.EndRow();

        // Row 4, Cell 1 – empty (to make it 2x2)
        builder.InsertCell();
        builder.EndRow();

        // End table
        builder.EndTable();

        // Save the document
        doc.Save(docPath);

        // 3. Load the document for extraction
        Document loadedDoc = new Document(docPath);

        // 4. Extract images that are inside tables
        List<string> extractedImagePaths = new List<string>();
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);
        int imageIndex = 0;

        foreach (Table tbl in tables)
        {
            NodeCollection shapes = tbl.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"image_{imageIndex}.png";
                    string imageFullPath = Path.Combine(tempImageFolder, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    extractedImagePaths.Add(imageFullPath);
                    imageIndex++;
                }
            }
        }

        // Validate extraction
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from tables.");

        // 5. Create a zip archive containing the extracted images
        using (FileStream zipToOpen = new FileStream(outputZipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (string imgPath in extractedImagePaths)
            {
                string entryName = Path.GetFileName(imgPath);
                archive.CreateEntryFromFile(imgPath, entryName);
            }
        }

        // Validate zip creation
        if (!File.Exists(outputZipPath))
            throw new InvalidOperationException("Zip archive was not created.");

        // Clean up temporary images
        Directory.Delete(tempImageFolder, true);
    }
}
