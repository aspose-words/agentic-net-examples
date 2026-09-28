using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDocsDir = Path.Combine(baseDir, "InputDocs");
        string imagesDir = Path.Combine(baseDir, "ExtractedImages");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDocsDir);
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files each containing the sample image
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputDocsDir, $"Doc{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // Prepare data for index
        var indexRows = new List<(string DocPath, string ImagePath)>();

        // Process each DOCX file
        foreach (string docFile in Directory.GetFiles(inputDocsDir, "*.docx"))
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_Image{imageIndex}.png";
                    string imagePath = Path.Combine(imagesDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    indexRows.Add((docFile, imagePath));
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (indexRows.Count == 0)
            throw new InvalidOperationException("No images were extracted from the documents.");

        // Create a simple CSV file with .xlsx extension as the index
        string excelPath = Path.Combine(outputDir, "ImageIndex.xlsx");
        CreateCsvIndex(excelPath, indexRows);

        // Validate Excel (CSV) file creation
        if (!File.Exists(excelPath))
            throw new InvalidOperationException("Failed to create the Excel index file.");
    }

    // Creates a simple white image with a black rectangle using Aspose.Drawing
    private static void CreateSampleImage(string path, int width, int height)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        try
        {
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            finally
            {
                g.Dispose();
            }

            bitmap.Save(path, ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }
    }

    // Creates a DOCX file and inserts the specified image
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document generated for image extraction: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    // Generates a CSV file (named with .xlsx extension) listing document and image paths
    private static void CreateCsvIndex(string filePath, List<(string DocPath, string ImagePath)> rows)
    {
        using (StreamWriter writer = new StreamWriter(filePath, false))
        {
            // Header
            writer.WriteLine("Document Path,Extracted Image Path");

            // Data rows
            foreach (var row in rows)
            {
                // Escape commas if needed
                string docPath = $"\"{row.DocPath}\"";
                string imgPath = $"\"{row.ImagePath}\"";
                writer.WriteLine($"{docPath},{imgPath}");
            }
        }
    }
}
