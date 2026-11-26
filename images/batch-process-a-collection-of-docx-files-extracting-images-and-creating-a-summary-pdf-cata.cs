using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Newtonsoft.Json;

public class ImageBatchProcessor
{
    public static void Main()
    {
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string imageFolder = Path.Combine(baseDir, "ExtractedImages");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(imageFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(baseDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Create sample DOCX files containing the sample image
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            CreateSampleDocumentWithImage(docPath, sampleImagePath, $"Sample document {i}");
        }

        // Extract images from each DOCX and collect info for the catalog
        var extractedInfo = new List<ExtractedImageInfo>();
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 1;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_Image{imageIndex}.png";
                    string imagePath = Path.Combine(imageFolder, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedInfo.Add(new ExtractedImageInfo
                    {
                        SourceDocument = Path.GetFileName(docPath),
                        ImageFileName = imageFileName,
                        ImagePath = imagePath
                    });
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted
        if (extractedInfo.Count == 0)
            throw new InvalidOperationException("No images were extracted from the DOCX files.");

        // Create a PDF catalog summarizing the extracted images
        Document catalog = new Document();
        DocumentBuilder builder = new DocumentBuilder(catalog);
        foreach (var info in extractedInfo)
        {
            builder.Writeln($"Source Document: {info.SourceDocument}");
            builder.Writeln($"Image: {info.ImageFileName}");
            builder.InsertImage(info.ImagePath);
            builder.Writeln(); // Add spacing
        }

        string catalogPath = Path.Combine(outputFolder, "ImageCatalog.pdf");
        catalog.Save(catalogPath, SaveFormat.Pdf);

        // Validate PDF creation
        if (!File.Exists(catalogPath))
            throw new InvalidOperationException("Failed to create the PDF catalog.");

        // Optional: write a JSON manifest of extracted images
        string manifestPath = Path.Combine(outputFolder, "manifest.json");
        string json = JsonConvert.SerializeObject(extractedInfo, Formatting.Indented);
        File.WriteAllText(manifestPath, json);
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle
                using (var pen = new Pen(Color.Black, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path);
        }
    }

    private static void CreateSampleDocumentWithImage(string docPath, string imagePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(title);
        builder.InsertParagraph();
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }

    private class ExtractedImageInfo
    {
        public string SourceDocument { get; set; }
        public string ImageFileName { get; set; }
        public string ImagePath { get; set; }
    }
}
