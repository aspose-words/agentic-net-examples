using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Base working directory for the demo
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImageExtractionDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string imageOutputDir = Path.Combine(baseDir, "ExtractedImages");
        string indexPath = Path.Combine(baseDir, "ImageIndex.xlsx"); // CSV with .xlsx extension

        // Ensure clean environment
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(imageOutputDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (input.png)
        // -------------------------------------------------
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create sample DOCX files each containing the image
        // -------------------------------------------------
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i}");
            builder.InsertImage(sampleImagePath);
            string docPath = Path.Combine(inputDir, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // -------------------------------------------------
        // 3. Batch process DOCX files: extract images and build index data
        // -------------------------------------------------
        var indexEntries = new List<(string DocumentPath, string ImagePath)>();
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx", SearchOption.TopDirectoryOnly);
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}.png";
                    string imagePath = Path.Combine(imageOutputDir, imageFileName);
                    // Save the image; format inferred from file extension
                    shape.ImageData.Save(imagePath);
                    indexEntries.Add((docFile, imagePath));
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (indexEntries.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // -------------------------------------------------
        // 4. Generate a CSV index and save with .xlsx extension
        // -------------------------------------------------
        using (StreamWriter writer = new StreamWriter(indexPath, false))
        {
            writer.WriteLine("DocumentPath,ImagePath");
            foreach (var entry in indexEntries)
            {
                // Escape commas by surrounding with quotes
                string docPathEscaped = $"\"{entry.DocumentPath}\"";
                string imgPathEscaped = $"\"{entry.ImagePath}\"";
                writer.WriteLine($"{docPathEscaped},{imgPathEscaped}");
            }
        }

        // Validate index file existence
        if (!File.Exists(indexPath))
            throw new FileNotFoundException("Failed to create the Excel index file.", indexPath);

        // -------------------------------------------------
        // 5. (Optional) Serialize index data to JSON using Newtonsoft.Json
        // -------------------------------------------------
        string jsonPath = Path.Combine(baseDir, "ImageIndex.json");
        string jsonContent = JsonConvert.SerializeObject(indexEntries, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonContent);
    }
}
