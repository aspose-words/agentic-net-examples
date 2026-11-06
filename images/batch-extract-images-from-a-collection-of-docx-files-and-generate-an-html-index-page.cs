using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchImageExtractor
{
    public static void Main()
    {
        // Define folders
        string baseFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string imageFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");
        string htmlPath = Path.Combine(Directory.GetCurrentDirectory(), "index.html");

        // Ensure clean environment
        if (Directory.Exists(baseFolder)) Directory.Delete(baseFolder, true);
        if (Directory.Exists(imageFolder)) Directory.Delete(imageFolder, true);
        Directory.CreateDirectory(baseFolder);
        Directory.CreateDirectory(imageFolder);

        // Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(Directory.GetCurrentDirectory(), "input.png");
        CreateSampleImage(sampleImagePath);

        // Create sample DOCX files containing the sample image
        int docCount = 3;
        for (int i = 1; i <= docCount; i++)
        {
            string docPath = Path.Combine(baseFolder, $"Document{i}.docx");
            CreateDocumentWithImage(docPath, sampleImagePath);
        }

        // Dictionary to hold document name and its extracted image file names
        var docImagesMap = new Dictionary<string, List<string>>();

        // Process each DOCX file
        foreach (string docFile in Directory.GetFiles(baseFolder, "*.docx"))
        {
            var extractedImages = ExtractImagesFromDocument(docFile, imageFolder);
            if (extractedImages.Count == 0)
                throw new Exception($"No images were extracted from {Path.GetFileName(docFile)}.");

            docImagesMap[Path.GetFileName(docFile)] = extractedImages;
        }

        // Generate HTML index page
        GenerateHtmlIndex(htmlPath, docImagesMap);
        if (!File.Exists(htmlPath))
            throw new Exception("HTML index page was not created.");

        // Validation: at least one image file should exist
        if (Directory.GetFiles(imageFolder).Length == 0)
            throw new Exception("No extracted images were saved.");
    }

    private static void CreateSampleImage(string filePath)
    {
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle
                using (Pen pen = new Pen(Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new Exception("Sample image was not created.");
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"This is a sample document: {Path.GetFileName(docPath)}");
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception($"Document {Path.GetFileName(docPath)} was not saved.");
    }

    private static List<string> ExtractImagesFromDocument(string docPath, string outputFolder)
    {
        var savedImages = new List<string>();
        Document doc = new Document(docPath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 1;
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                // Determine file extension based on image type
                string extension = "." + shape.ImageData.ImageType.ToString().ToLower();

                string imageFileName = $"{Path.GetFileNameWithoutExtension(docPath)}_image{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputFolder, imageFileName);
                shape.ImageData.Save(imageFullPath);
                if (!File.Exists(imageFullPath))
                    throw new Exception($"Failed to save extracted image {imageFileName}.");

                savedImages.Add(imageFileName);
                imageIndex++;
            }
        }

        return savedImages;
    }

    private static void GenerateHtmlIndex(string htmlPath, Dictionary<string, List<string>> docImagesMap)
    {
        using (StreamWriter writer = new StreamWriter(htmlPath, false))
        {
            writer.WriteLine("<!DOCTYPE html>");
            writer.WriteLine("<html>");
            writer.WriteLine("<head><meta charset=\"UTF-8\"><title>Extracted Images Index</title></head>");
            writer.WriteLine("<body>");
            writer.WriteLine("<h1>Extracted Images Index</h1>");

            foreach (var entry in docImagesMap)
            {
                writer.WriteLine($"<h2>{entry.Key}</h2>");
                foreach (string imgFile in entry.Value)
                {
                    string imgSrc = Path.Combine("ExtractedImages", imgFile).Replace("\\", "/");
                    writer.WriteLine($"<img src=\"{imgSrc}\" alt=\"{imgFile}\" style=\"max-width:300px; margin:5px;\"/>");
                }
            }

            writer.WriteLine("</body>");
            writer.WriteLine("</html>");
        }
    }
}
